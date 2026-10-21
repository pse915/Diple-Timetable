import os
from pathlib import Path
import tempfile
import unittest

from desktop.core import LocalStore, parse_school_workbook, ValidationError, school_catalog
from desktop.xlsx_io import read_xlsx, write_xlsx, WorkbookError

TEMPLATE=Path(__file__).parents[1]/'templates'/'학교_시간표_양식.xlsx'

class SchoolTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.before=os.environ.get('SCHOOL_TIMETABLE_HOME')
        os.environ['SCHOOL_TIMETABLE_HOME']=self.tmp.name
        self.sample=Path(self.tmp.name)/'test_school.xlsx'
        test_data=read_xlsx(TEMPLATE)
        test_data['학교정보'][1][0]='테스트용 학교'
        write_xlsx(self.sample,test_data)
        self.input=parse_school_workbook(self.sample)
        self.store=LocalStore.for_info(self.input)
        self.store.import_workbook(self.sample)

    def tearDown(self):
        if self.before is None: os.environ.pop('SCHOOL_TIMETABLE_HOME',None)
        else: os.environ['SCHOOL_TIMETABLE_HOME']=self.before
        self.tmp.cleanup()

    def test_blank_template_requires_school_name(self):
        with self.assertRaises(ValidationError):
            parse_school_workbook(TEMPLATE)

    def test_import_and_catalog(self):
        data=self.store.snapshot()
        self.assertEqual(data['school'],'테스트용 학교')
        self.assertEqual(len(data['slots']),75)
        self.assertEqual(len(data['classes']),3)
        self.assertEqual(len(school_catalog()),1)

    def test_edit_success_and_history(self):
        self.store.edit_slot('1-1','월',1,'프로젝트수업','','기술실')
        data=self.store.snapshot()
        result=next(s for s in data['slots'] if s['class']=='1-1' and s['day']=='월' and s['period']==1)
        self.assertEqual(result['subject'],'프로젝트수업')
        self.assertEqual(self.store.recent_history()[0]['action'],'EDIT')

    def test_reject_unknown_teacher_without_writes(self):
        before=self.store.snapshot()
        with self.assertRaises(ValidationError):
            self.store.edit_slot('1-1','월',1,'수학','없음선생님','1-1')
        self.assertEqual(before['slots'],self.store.snapshot()['slots'])

    def test_reject_teacher_overlap(self):
        before=self.store.snapshot()
        # 1-2 월 1교시 교사2 assigned; attempt same teacher in 1-1.
        with self.assertRaises(ValidationError):
            self.store.edit_slot('1-1','월',1,'영어','교사2','1-1')
        self.assertEqual(before['slots'],self.store.snapshot()['slots'])

    def test_swap_valid_and_logged(self):
        a=('1-1','월',1); b=('1-1','월',4)
        pre={ (r['class'],r['day'],r['period']):r for r in self.store.snapshot()['slots']}
        self.store.swap(a,b)
        post={ (r['class'],r['day'],r['period']):r for r in self.store.snapshot()['slots']}
        self.assertEqual(post[a]['subject'],pre[b]['subject'])
        self.assertEqual(post[b]['teacher'],pre[a]['teacher'])
        self.assertEqual(self.store.recent_history()[0]['action'],'SWAP')

    def test_swap_collision_rolls_back(self):
        before=self.store.snapshot()
        with self.assertRaises(ValidationError):
            self.store.swap(('1-1','월',1),('1-1','월',2))
        self.assertEqual(before['slots'],self.store.snapshot()['slots'])

    def test_export_roundtrip(self):
        self.store.edit_slot('1-1','월',1,'=SUM(1,2)','','')
        exported=Path(self.tmp.name)/'export.xlsx'
        self.store.export_workbook(exported)
        wb=read_xlsx(exported)
        row=next(r for r in wb['시간표'][1:] if r[:3]==['1-1','월','1'])
        self.assertEqual(row[3],'=SUM(1,2)')
        self.assertEqual(len(wb['변경이력']),3)
        self.assertEqual(len(parse_school_workbook(exported)['slots']),75)

    def test_original_file_backup_on_reimport(self):
        self.store.edit_slot('1-1','월',1,'재입력전','','')
        self.store.import_workbook(self.sample,replace=True)
        backups=list((Path(self.tmp.name)/'SchoolTimetableDesktop'/'backups').glob('*.sqlite3'))
        self.assertEqual(len(backups),1)
        self.assertNotEqual(self.store.snapshot()['slots'][0]['subject'],'재입력전')

    def test_reject_missing_teacher_sheet(self):
        bad=Path(self.tmp.name)/'broken.xlsx'
        write_xlsx(bad,{'학교정보':[['학교명','학년도'],['테스트','2026']]})
        with self.assertRaises(ValidationError): parse_school_workbook(bad)

    def test_other_school_separated(self):
        other=Path(self.tmp.name)/'other.xlsx'
        wb=read_xlsx(self.sample)
        wb['학교정보'][1][0]='다른 테스트용 학교'
        write_xlsx(other,wb)
        new=LocalStore.for_info(parse_school_workbook(other));new.import_workbook(other)
        self.assertNotEqual(new.path,self.store.path)
        new.edit_slot('1-1','월',1,'다른학교수업','','')
        self.assertNotEqual(self.store.snapshot()['slots'][0]['subject'],'다른학교수업')

if __name__=='__main__': unittest.main()
