"""Local, per-school SQLite timetable model; never contacts Google or the internet."""
from __future__ import annotations
import contextlib
from dataclasses import dataclass
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
from typing import Optional
from .xlsx_io import read_xlsx, write_xlsx, WorkbookError

DAY_ORDER = ['월','화','수','목','금','토','일']
SLOT_HEADERS = ['학급','요일','교시','교과','교사','교실']

class ValidationError(ValueError):
    pass


def app_home() -> Path:
    p = Path(os.environ.get('SCHOOL_TIMETABLE_HOME') or os.environ.get('LOCALAPPDATA') or (Path.home()/'.local'/'share'))/'SchoolTimetableDesktop'
    p.mkdir(parents=True, exist_ok=True)
    return p


def _table(data, name, columns, optional=False):
    rows = data.get(name)
    if not rows:
        if optional: return []
        raise ValidationError(f'필수 시트 [{name}]가 없습니다.')
    headers = [str(x).strip() for x in rows[0]]
    if any(h not in headers for h in columns):
        raise ValidationError(f'[{name}] 시트의 필수 열: {", ".join(columns)}')
    idx=[headers.index(c) for c in columns]
    return [dict(zip(columns,[str(row[i]).strip() if i < len(row) else '' for i in idx])) for row in rows[1:] if any(str(v).strip() for v in row)]


def _period(v):
    try:
        if isinstance(v,str) and v.endswith('.0'): v = v[:-2]
        n=int(v)
    except (TypeError,ValueError) as e:
        raise ValidationError(f'교시는 1~12 사이 정수여야 합니다: {v}') from e
    if not 1 <= n <= 12: raise ValidationError(f'교시 범위 오류: {v}')
    return n


def parse_school_workbook(path):
    data = read_xlsx(path)
    info = _table(data,'학교정보',['학교명','학년도'])
    if not info or not info[0]['학교명']:
        raise ValidationError('학교정보 시트에 학교명이 필요합니다.')
    school=info[0]['학교명'].strip()
    if school in ('학교명을 입력하세요', '학교명 입력', '학교명'):
        raise ValidationError('학교정보 시트의 A2 셀에 실제 학교명을 입력하세요.')
    year=info[0]['학년도'].strip()
    if not re.fullmatch(r'20\d{2}',year): raise ValidationError('학년도는 2026처럼 4자리 연도여야 합니다.')
    teachers = _table(data,'교사',['교사명','담당교과'])
    classes = _table(data,'학급',['학급'])
    slots = _table(data,'시간표',SLOT_HEADERS)
    teacher_names = [r['교사명'] for r in teachers if r['교사명']]
    class_names = [r['학급'] for r in classes if r['학급']]
    if len(teacher_names) != len(set(teacher_names)) or not teacher_names:
        raise ValidationError('교사명이 중복되었거나 교사 목록이 비어 있습니다.')
    if len(class_names) != len(set(class_names)) or not class_names:
        raise ValidationError('학급명이 중복되었거나 학급 목록이 비어 있습니다.')
    used_slots = set(); occupied = set(); parsed=[]
    for i,r in enumerate(slots,2):
        cls,day = r['학급'],r['요일']
        period = _period(r['교시'])
        if cls not in class_names or day not in DAY_ORDER:
            raise ValidationError(f'시간표 {i}행: 학급 또는 요일이 유효하지 않습니다.')
        key = (cls,day,period)
        if key in used_slots: raise ValidationError(f'시간표에 중복된 학급·요일·교시: {key}')
        used_slots.add(key)
        teacher=r['교사']; subject=r['교과']; room=r['교실']
        if teacher and teacher not in teacher_names:
            raise ValidationError(f'시간표 {i}행: 등록되지 않은 교사 {teacher}')
        if teacher and not subject: raise ValidationError(f'시간표 {i}행: 교과 없이 교사만 지정되었습니다.')
        if teacher:
            teachkey=(teacher,day,period)
            if teachkey in occupied: raise ValidationError(f'중복 배정 교사: {teachkey}')
            occupied.add(teachkey)
        parsed.append({'class':cls,'day':day,'period':period,'subject':subject,'teacher':teacher,'room':room})
    return {'school':school,'year':year,'teachers':[{'name':r['교사명'],'subject':r['담당교과']} for r in teachers], 'classes':class_names,'slots':parsed}


def school_key(info):
    return hashlib.sha256((info['school']+'|'+info['year']).encode('utf-8')).hexdigest()[:18]


@dataclass
class LocalStore:
    path: Path

    @contextlib.contextmanager
    def connect(self):
        c=sqlite3.connect(str(self.path), timeout=15)
        c.row_factory=sqlite3.Row
        c.execute('PRAGMA foreign_keys=ON')
        c.execute('PRAGMA busy_timeout=15000')
        try:
            with c:
                yield c
        finally:
            c.close()

    @classmethod
    def for_info(cls,info):
        path=app_home()/'schools'/school_key(info)/'data.sqlite3'
        path.parent.mkdir(parents=True,exist_ok=True)
        return cls(path)

    @classmethod
    def for_key(cls,key):
        if not re.fullmatch(r'[0-9a-f]{18}',key): raise ValidationError('잘못된 학교 키')
        return cls(app_home()/'schools'/key/'data.sqlite3')

    def _ensure(self,c):
        c.executescript('''
        CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS teachers(name TEXT PRIMARY KEY, subject TEXT NOT NULL DEFAULT '');
        CREATE TABLE IF NOT EXISTS classes(name TEXT PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS slots(class TEXT NOT NULL, day TEXT NOT NULL,period INTEGER NOT NULL,
          subject TEXT NOT NULL DEFAULT '', teacher TEXT NOT NULL DEFAULT '', room TEXT NOT NULL DEFAULT '',
          PRIMARY KEY(class,day,period),FOREIGN KEY(class) REFERENCES classes(name));
        CREATE TABLE IF NOT EXISTS history(id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, action TEXT NOT NULL,
          before_json TEXT NOT NULL, after_json TEXT NOT NULL,note TEXT NOT NULL DEFAULT '');
        ''')

    def import_workbook(self, path, replace=False):
        info=parse_school_workbook(path)
        if self.path.resolve()!=self.for_info(info).path.resolve():
            raise ValidationError('선택한 학교 데이터베이스가 엑셀의 학교정보와 일치하지 않습니다.')
        self.path.parent.mkdir(parents=True,exist_ok=True)
        if self.path.exists() and not replace:
            raise ValidationError('기존 학교 데이터가 있습니다. 덮어쓰기 승인이 필요합니다.')
        # Build a complete DB in a temporary file and atomically replace; preserve prior DB by backup.
        tmp=self.path.with_suffix('.import.tmp')
        if tmp.exists(): tmp.unlink()
        try:
            with contextlib.closing(sqlite3.connect(tmp)) as c, c:
                self._ensure(c)
                c.executemany('INSERT INTO meta(key,value) VALUES (?,?)', [('학교명',info['school']),('학년도',info['year'])])
                c.executemany('INSERT INTO teachers VALUES (?,?)',[(x['name'],x['subject']) for x in info['teachers']])
                c.executemany('INSERT INTO classes VALUES (?)',[(x,) for x in info['classes']])
                c.executemany('INSERT INTO slots VALUES (?,?,?,?,?,?)',[(s['class'],s['day'],s['period'],s['subject'],s['teacher'],s['room']) for s in info['slots']])
                c.execute('INSERT INTO history(at,action,before_json,after_json,note) VALUES (?,?,?,?,?)',
                    (datetime.now().isoformat(timespec='seconds'),'IMPORT','[]','[]',str(Path(path).name)))
            if self.path.exists():
                backups=app_home()/'backups'; backups.mkdir(exist_ok=True)
                timestamp=datetime.now().strftime('%Y%m%d-%H%M%S-%f')
                shutil.copy2(self.path,backups/f'{self.path.parent.name}-{timestamp}.sqlite3')
            os.replace(tmp,self.path)
        finally:
            if tmp.exists(): tmp.unlink()
        return info

    def snapshot(self):
        with self.connect() as c:
            meta={r['key']:r['value'] for r in c.execute('SELECT * FROM meta')}
            classes=[r['name'] for r in c.execute('SELECT name FROM classes ORDER BY name')]
            teachers=[dict(r) for r in c.execute('SELECT * FROM teachers ORDER BY name')]
            slots=[dict(r) for r in c.execute('SELECT * FROM slots ORDER BY class,period,day')]
            return {'school':meta.get('학교명',''),'year':meta.get('학년도',''), 'classes':classes,'teachers':teachers,'slots':slots}

    @staticmethod
    def _check_collision(c, row_updates):
        # Test the *final* timetable state, not the intermediate one.
        slots={(r['class'],r['day'],r['period']):dict(r) for r in c.execute('SELECT * FROM slots')}
        for r in row_updates:
            slots[(r['class'],r['day'],r['period'])]=r
        occupancy={}
        for r in slots.values():
            if r['teacher']:
                tag=(r['teacher'],r['day'],r['period'])
                previous=occupancy.get(tag)
                if previous and previous!=r['class']:
                    raise ValidationError(f'교사 시간 중복: {r["teacher"]} / {r["day"]} / {r["period"]}교시 ({previous}, {r["class"]})')
                occupancy[tag]=r['class']

    def _update(self, updates, action, note):
        with self.connect() as c:
            c.execute('BEGIN IMMEDIATE')
            teachers={r['name'] for r in c.execute('SELECT name FROM teachers')}
            classes={r['name'] for r in c.execute('SELECT name FROM classes')}
            before=[]
            for item in updates:
                if item['class'] not in classes or item['day'] not in DAY_ORDER or _period(item['period'])!=item['period']:
                    raise ValidationError('학급/요일/교시 오류')
                if item['teacher'] and item['teacher'] not in teachers:
                    raise ValidationError('교사 명단에 없는 교사: '+item['teacher'])
                if item['teacher'] and not item['subject']:
                    raise ValidationError('교과가 없는 교사 배정은 허용되지 않습니다.')
                r=c.execute('SELECT * FROM slots WHERE class=? AND day=? AND period=?',
                   (item['class'],item['day'],item['period'])).fetchone()
                before.append(dict(r) if r else {'class':item['class'],'day':item['day'],'period':item['period'], 'subject':'','teacher':'','room':''})
            self._check_collision(c,updates)
            for item in updates:
                c.execute('INSERT INTO slots(class,day,period,subject,teacher,room) VALUES(?,?,?,?,?,?) ON CONFLICT(class,day,period) DO UPDATE SET subject=excluded.subject,teacher=excluded.teacher,room=excluded.room',
                   tuple(item[k] for k in ['class','day','period','subject','teacher','room']))
            c.execute('INSERT INTO history(at,action,before_json,after_json,note) VALUES (?,?,?,?,?)',
                 (datetime.now().isoformat(timespec='seconds'),action,json.dumps(before,ensure_ascii=False),json.dumps(updates,ensure_ascii=False),note))

    def edit_slot(self,cls,day,period,subject,teacher,room):
        return self._update([{'class':cls,'day':day,'period':period,'subject':subject.strip(),'teacher':teacher.strip(),'room':room.strip()}],'EDIT',f'{cls} {day} {period}교시 수정')

    def swap(self,a,b):
        if a==b: raise ValidationError('서로 다른 두 수업을 선택하세요.')
        with self.connect() as c:
            def get(k):
                row=c.execute('SELECT * FROM slots WHERE class=? AND day=? AND period=?',k).fetchone()
                return dict(row) if row else {'class':k[0],'day':k[1],'period':k[2],'subject':'','teacher':'','room':''}
            av,bv=get(a),get(b)
        updates=[]
        for target,source in [(a,bv),(b,av)]:
            updates.append({'class':target[0],'day':target[1],'period':target[2],'subject':source['subject'],'teacher':source['teacher'],'room':source['room']})
        self._update(updates,'SWAP',f'{a} ↔ {b}')

    def recent_history(self,count=200):
        with self.connect() as c:
            return [dict(r) for r in c.execute('SELECT id,at,action,note FROM history ORDER BY id DESC LIMIT ?',(count,))]

    def export_workbook(self,path):
        info=self.snapshot()
        rows={
            '학교정보':[['학교명','학년도'],[info['school'],info['year']]],
            '교사':[['교사명','담당교과']]+[[x['name'],x['subject']] for x in info['teachers']],
            '학급':[['학급']]+[[x] for x in info['classes']],
            '시간표':[SLOT_HEADERS]+[[s['class'],s['day'],str(s['period']),s['subject'],s['teacher'],s['room']] for s in info['slots']],
            '변경이력':[['번호','시각','작업','내용']]+[[str(h['id']),h['at'],h['action'],h['note']] for h in reversed(self.recent_history(5000))]
        }
        write_xlsx(path,rows)


def school_catalog():
    schools=app_home()/'schools'
    found=[]
    if not schools.exists(): return []
    for path in sorted(schools.glob('*/data.sqlite3')):
        try:
            snap=LocalStore(path).snapshot()
            if snap['school']:
                found.append((path.parent.name, f"{snap['school']} ({snap['year']})"))
        except (sqlite3.DatabaseError,OSError): continue
    return found
