"""Single-machine offline Windows timetable manager.
Run with `python main.py` while developing, or package with PyInstaller on Windows.
"""
from __future__ import annotations
import json
from pathlib import Path
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from desktop.core import DAY_ORDER, LocalStore, app_home, parse_school_workbook, school_catalog, ValidationError
from desktop.xlsx_io import WorkbookError

DAYS=DAY_ORDER[:5]
TITLE='Diple Timetable'
BG='#F4F7FB'; TEXT='#18304A'; BLUE='#275EAC'


def resource_path(relative):
    return Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent))/relative


class EditDialog(tk.Toplevel):
    def __init__(self,parent,initial,teachers):
        super().__init__(parent)
        self.result=None
        self.title('수업 상세 편집')
        self.resizable(False,False)
        self.transient(parent)
        self.grab_set()
        self.columnconfigure(1,weight=1)
        ttk.Label(self,text=f'{initial["class"]}  ·  {initial["day"]}요일  {initial["period"]}교시',
                  font=('맑은 고딕',12,'bold')).grid(row=0,column=0,columnspan=2,padx=22,pady=(20,12),sticky='w')
        self.subject=tk.StringVar(value=initial['subject']); self.teacher=tk.StringVar(value=initial['teacher']); self.room=tk.StringVar(value=initial['room'])
        for row,(label,var) in enumerate([('교과',self.subject),('교사',self.teacher),('교실',self.room)],1):
            ttk.Label(self,text=label).grid(row=row,column=0,padx=20,pady=8,sticky='w')
            if row==2:
                control=ttk.Combobox(self,textvariable=var,values=['']+teachers,width=30)
            else:
                control=ttk.Entry(self,textvariable=var,width=33)
            control.grid(row=row,column=1,padx=(0,20),pady=8,sticky='ew')
        actions=ttk.Frame(self)
        actions.grid(row=4,column=0,columnspan=2,pady=18,padx=20,sticky='e')
        ttk.Button(actions,text='취소',command=self.destroy).pack(side='right',padx=6)
        ttk.Button(actions,text='저장',command=self.submit).pack(side='right')
        self.bind('<Escape>',lambda _e:self.destroy())
        self.bind('<Return>',lambda _e:self.submit())
        self.wait_visibility()
        self.focus_set()

    def submit(self):
        self.result=(self.subject.get(),self.teacher.get(),self.room.get())
        self.destroy()


class TimetableApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITLE)
        self.minsize(920,630)
        self.geometry('1120x760')
        self.configure(background=BG)
        style=ttk.Style(self)
        try: style.theme_use('clam')
        except tk.TclError: pass
        style.configure('Header.TLabel',font=('맑은 고딕',19,'bold'),foreground=TEXT,background=BG)
        style.configure('Info.TLabel',font=('맑은 고딕',10),foreground='#52677A',background=BG)
        style.configure('TButton',padding=(11,8))
        style.configure('Treeview',font=('맑은 고딕',10),rowheight=70)
        style.configure('Treeview.Heading',font=('맑은 고딕',10,'bold'),padding=(5,11))
        self.store=None; self.slots=[]; self.selected=[]; self.periods=[]; self.display={} ; self.catalog={}
        self.school_choice=tk.StringVar()
        self.mode=tk.StringVar(value='학급별')
        self.subject=tk.StringVar()
        self.status=tk.StringVar(value='엑셀 파일을 불러와 주세요.')
        self._build()
        self._load_saved_school()

    def _build(self):
        wrapper=ttk.Frame(self,padding=22)
        wrapper.pack(fill='both',expand=True)
        top=ttk.Frame(wrapper)
        top.pack(fill='x')
        ttk.Label(top,text='Diple Timetable',style='Header.TLabel').pack(side='left')
        ttk.Label(top,text='100% 로컬 저장 · Google 연동 없음',style='Info.TLabel').pack(side='right')
        sub=ttk.Frame(wrapper)
        sub.pack(fill='x',pady=(14,15))
        ttk.Button(sub,text='엑셀 불러오기',command=self.import_excel).pack(side='left',padx=(0,7))
        ttk.Button(sub,text='엑셀 내보내기',command=self.export_excel).pack(side='left',padx=7)
        ttk.Button(sub,text='예시 템플릿 저장',command=self.save_template).pack(side='left',padx=7)
        ttk.Button(sub,text='변경 이력',command=self.show_history).pack(side='left',padx=7)
        ttk.Button(sub,text='프로그램 정보',command=self.show_info).pack(side='right')
        group=ttk.LabelFrame(wrapper,text='학교와 시간표 선택',padding=(14,12))
        group.pack(fill='x',pady=(0,13))
        ttk.Label(group,text='학교').grid(row=0,column=0,sticky='w',padx=5)
        self.schools=ttk.Combobox(group,textvariable=self.school_choice,state='readonly',width=26)
        self.schools.grid(row=0,column=1,sticky='ew',padx=8)
        self.schools.bind('<<ComboboxSelected>>',self.change_school)
        for i,name in enumerate(['학급별','교사별']):
            ttk.Radiobutton(group,text=name,variable=self.mode,value=name,command=self.switch_mode).grid(row=0,column=2+i,padx=8)
        ttk.Label(group,text='조회 대상').grid(row=0,column=4,sticky='w',padx=(18,5))
        self.targets=ttk.Combobox(group,textvariable=self.subject,state='readonly',width=18)
        self.targets.grid(row=0,column=5,padx=8)
        self.targets.bind('<<ComboboxSelected>>',lambda _e:self.refresh_table())
        group.columnconfigure(1,weight=1); group.columnconfigure(5,weight=1)
        ttk.Label(wrapper,text='시간표  |  학급별 화면에서 첫 번째 수업과 두 번째 수업을 차례로 선택한 뒤 맞교환을 실행합니다.',style='Info.TLabel').pack(anchor='w',pady=(0,8))
        treebox=ttk.Frame(wrapper)
        treebox.pack(fill='both',expand=True)
        cols=['period']+DAYS
        self.tree=ttk.Treeview(treebox,columns=cols,show='headings',selectmode='browse')
        for k,name in zip(cols,['교시']+[d+'요일' for d in DAYS]):
            self.tree.heading(k,text=name)
            self.tree.column(k,width=72 if k=='period' else 185,anchor='center',stretch=k!='period')
        vbar=ttk.Scrollbar(treebox,orient='vertical',command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)
        self.tree.pack(side='left',fill='both',expand=True)
        vbar.pack(side='left',fill='y')
        self.tree.bind('<ButtonRelease-1>',self.on_cell_click)
        self.tree.bind('<Double-Button-1>',self.on_double_click)
        bottom=ttk.Frame(wrapper)
        bottom.pack(fill='x',pady=(12,4))
        ttk.Button(bottom,text='선택한 수업 편집',command=self.edit_selected).pack(side='left',padx=(0,8))
        ttk.Button(bottom,text='선택한 두 수업 맞교환',command=self.swap_selected).pack(side='left',padx=8)
        ttk.Button(bottom,text='선택 해제',command=self.reset_selected).pack(side='left',padx=8)
        ttk.Label(wrapper,textvariable=self.status,style='Info.TLabel',wraplength=960).pack(anchor='w',pady=(8,0))

    def _load_saved_school(self):
        schools=school_catalog()
        self.catalog={label:key for key,label in schools}
        self.schools.configure(values=list(self.catalog))
        settings=app_home()/'settings.json'
        last=None
        if settings.exists():
            try: last=json.loads(settings.read_text('utf-8')).get('active_school')
            except (OSError,ValueError): pass
        target=next((label for key,label in schools if key==last),None)
        if not target and schools: target=schools[0][1]
        if target:
            self.school_choice.set(target)
            self.change_school()
        else: self.status.set('처음 사용하신다면 [예시 템플릿 저장] 후 내용을 수정해 [엑셀 불러오기]를 눌러 주세요.')

    def change_school(self,_evt=None):
        key=self.catalog.get(self.school_choice.get())
        if not key: return
        store=LocalStore.for_key(key)
        if not store.path.exists(): return
        self.store=store
        (app_home()/'settings.json').write_text(json.dumps({'active_school':key}),encoding='utf-8')
        self.switch_mode()

    def switch_mode(self):
        self.reset_selected()
        if self.store:
            info=self.store.snapshot()
            choices=info['classes'] if self.mode.get()=='학급별' else [t['name'] for t in info['teachers']]
            self.targets.configure(values=choices)
            self.subject.set(choices[0] if choices else '')
            self.refresh_table()

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        self.display={}
        if not self.store: return
        info=self.store.snapshot(); self.slots=info['slots']
        target=self.subject.get()
        if not target: return
        slotmap={ (s['class'],s['day'],s['period']):s for s in info['slots']}
        self.periods=list(range(1,max(7,max([s['period'] for s in self.slots]+[0]))+1))
        for period in self.periods:
            cells=[]
            for day in DAYS:
                s=slotmap.get((target,day,period)) if self.mode.get()=='학급별' else next((x for x in info['slots'] if x['teacher']==target and x['day']==day and x['period']==period),None)
                if s:
                    label=f"{s['subject']}\n{ s['teacher'] if self.mode.get()=='학급별' else s['class']}".strip()
                    self.display[(day,period)]=s
                else:
                    label='—'
                    if self.mode.get()=='학급별':
                        self.display[(day,period)]={'class':target,'day':day,'period':period,'subject':'','teacher':'','room':''}
                cells.append(label)
            self.tree.insert('', 'end', iid=str(period),values=[f'{period}교시']+cells)
        self.status.set(f"{info['school']} · {info['year']}학년도 | {self.mode.get()} {target} | 전체 {len(info['slots'])}개 수업 | 저장 위치: {self.store.path}")

    def _clicked(self,event):
        if self.mode.get()!='학급별': return None
        iid=self.tree.identify_row(event.y)
        column=self.tree.identify_column(event.x)
        if not iid or column not in ['#1','#2','#3','#4','#5','#6'] or column=='#1': return None
        return (DAYS[int(column[1:])-2],int(iid))

    def on_cell_click(self,event):
        cell=self._clicked(event)
        if cell is None: return
        if cell in self.selected: self.selected.remove(cell)
        elif len(self.selected)>=2: self.selected=[cell]
        else: self.selected.append(cell)
        self.status.set('선택: '+ ' / '.join(f'{day} {period}교시' for day,period in self.selected))

    def on_double_click(self,event):
        cell=self._clicked(event)
        if cell is not None:
            self.selected=[cell]; self.edit_selected()

    def reset_selected(self):
        self.selected=[]
        self.status.set('선택이 해제되었습니다.')

    def edit_selected(self):
        if not self.store or self.mode.get()!='학급별' or len(self.selected)!=1:
            messagebox.showinfo('수업 선택','학급별 시간표에서 수업 칸 하나를 선택해 주세요.'); return
        day,period=self.selected[0]; cls=self.subject.get()
        info=self.store.snapshot()
        current=next((x for x in info['slots'] if x['class']==cls and x['day']==day and x['period']==period),
            {'class':cls,'day':day,'period':period,'subject':'','teacher':'','room':''})
        popup=EditDialog(self,current,[t['name'] for t in info['teachers']])
        self.wait_window(popup)
        if popup.result is None: return
        try:
            self.store.edit_slot(cls,day,period,*popup.result)
            self.refresh_table(); self.reset_selected()
        except (ValidationError,ValueError) as e: messagebox.showerror('변경 불가',str(e))

    def swap_selected(self):
        if not self.store or self.mode.get()!='학급별' or len(self.selected)!=2:
            messagebox.showinfo('맞교환','학급별 시간표에서 교환할 두 수업을 순서대로 클릭하세요.'); return
        cls=self.subject.get()
        a,b=[(cls,day,period) for day,period in self.selected]
        if not messagebox.askyesno('수업 맞교환',f'{cls} {a[1]} {a[2]}교시 ↔ {b[1]} {b[2]}교시\n변경하시겠습니까?'):
            return
        try:
            self.store.swap(a,b); self.refresh_table(); self.reset_selected()
            messagebox.showinfo('완료','수업 맞교환을 저장했습니다.')
        except (ValidationError,ValueError) as e: messagebox.showerror('맞교환 불가',str(e))

    def import_excel(self):
        filename=filedialog.askopenfilename(title='시간표 엑셀 선택',filetypes=[('Excel (*.xlsx)','*.xlsx')])
        if not filename: return
        try:
            info=parse_school_workbook(filename)
            store=LocalStore.for_info(info)
            n=len(info['slots'])
            if store.path.exists():
                ask=f'{info["school"]} ({info["year"]}) 데이터가 이미 있습니다.\n현재 데이터베이스는 자동 백업하고, 엑셀 데이터로 교체합니다.\n기존 변경 이력은 새 데이터에 이어지지 않습니다. 계속할까요?'
            else:
                ask=f'{info["school"]} ({info["year"]})\n교사 {len(info["teachers"])}명 · 학급 {len(info["classes"])}개 · 수업 {n}개를 저장할까요?'
            if not messagebox.askyesno('엑셀 가져오기 확인',ask): return
            store.import_workbook(filename,replace=True)
            self._load_saved_school()
            target=next((label for label,key in self.catalog.items() if key==store.path.parent.name),None)
            if target: self.school_choice.set(target); self.change_school()
            messagebox.showinfo('가져오기 완료','로컬 데이터베이스에 안전하게 저장했습니다. 기존 데이터가 있으면 자동 백업했습니다.')
        except (ValidationError,WorkbookError,OSError,ValueError) as e:
            messagebox.showerror('엑셀 불러오기 실패',str(e))

    def export_excel(self):
        if not self.store:
            messagebox.showinfo('학교 선택','먼저 학교 엑셀 파일을 가져오세요.'); return
        snap=self.store.snapshot()
        path=filedialog.asksaveasfilename(title='시간표 엑셀 내보내기',defaultextension='.xlsx',initialfile=f'{snap["school"]}_{snap["year"]}_시간표.xlsx',filetypes=[('Excel (*.xlsx)','*.xlsx')])
        if not path:return
        try: self.store.export_workbook(path); messagebox.showinfo('내보내기 완료',f'저장했습니다:\n{path}')
        except (OSError,WorkbookError) as e: messagebox.showerror('내보내기 실패',str(e))

    def save_template(self):
        source=resource_path('templates/학교_시간표_양식.xlsx')
        if not source.exists(): messagebox.showerror('파일 없음','템플릿 파일을 찾지 못했습니다.'); return
        path=filedialog.asksaveasfilename(title='엑셀 템플릿 저장',defaultextension='.xlsx',initialfile='학교_시간표_양식.xlsx',filetypes=[('Excel (*.xlsx)','*.xlsx')])
        if path:
            try: shutil.copyfile(source,path); messagebox.showinfo('완료','템플릿을 저장했습니다.')
            except OSError as e: messagebox.showerror('오류',str(e))

    def show_history(self):
        if not self.store:return messagebox.showinfo('안내','먼저 학교를 불러오세요.')
        top=tk.Toplevel(self); top.title('수업 변경 이력'); top.geometry('850x440')
        tree=ttk.Treeview(top,columns=('id','at','act','note'),show='headings')
        for key,label,width in [('id','번호',65),('at','변경 시각',190),('act','구분',90),('note','상세 내용',450)]:
            tree.heading(key,text=label); tree.column(key,width=width,stretch=key=='note')
        tree.pack(fill='both',expand=True,padx=14,pady=14)
        for h in self.store.recent_history():tree.insert('', 'end',values=(h['id'],h['at'],h['action'],h['note']))

    def show_info(self):
        messagebox.showinfo('Diple Timetable',
            '학교별 Windows 단독 PC 로컬 시간표 프로그램 (초기 버전)\n\n'
            '· Google Sheets/Streamlit/API 호출 없음\n· 로컬 SQLite 자동 저장\n· .xlsx 입출력과 기본 수업 편집/맞교환\n'
            '· 원본 Streamlit 모든 기능의 100% 이식본은 아닙니다.\n\n저장 폴더:\n'+str(app_home()))


if __name__=='__main__':
    try:
        TimetableApp().mainloop()
    except Exception as exc:
        # Uncaught GUI errors should never silently become a white/empty screen.
        try: messagebox.showerror('예기치 않은 오류',str(exc))
        except Exception: pass
        raise
