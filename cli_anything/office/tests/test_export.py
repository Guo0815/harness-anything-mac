from unittest.mock import Mock
import pytest
from cli_anything.office.core import export as e
from cli_anything.office.utils import office_backend as b

@pytest.mark.parametrize('preset,formats',[('docx',['odt','docx']),('pdf',['odt','docx','pdf'])])
def test_pipeline(monkeypatch,tmp_path,preset,formats):
 monkeypatch.setattr(e,'COM_AVAILABLE',False)
 monkeypatch.setattr(b,'get_backend',lambda:'libreoffice-headless')
 calls=[]
 def convert(source,fmt,output_path=None,**kwargs):
  calls.append((source,fmt,output_path));return {'output':output_path,'format':fmt}
 monkeypatch.setattr(b,'convert_to',convert)
 e.export({'type':'writer','content':[]},str(tmp_path/'out'),preset)
 assert [c[1] for c in calls]==formats
 if preset=='pdf':assert calls[-1][0]==calls[-2][2]

@pytest.mark.parametrize('preset',['docx','pdf'])
def test_extension_overwrite_guard(monkeypatch,tmp_path,preset):
 monkeypatch.setattr(e,'COM_AVAILABLE',False);monkeypatch.setattr(b,'get_backend',lambda:'libreoffice-headless')
 p=tmp_path/('out.'+preset);p.write_bytes(b'keep')
 with pytest.raises(FileExistsError):e.export({'type':'writer'},str(tmp_path/'out'),preset)
 assert p.read_bytes()==b'keep'

def test_wrong_type(monkeypatch,tmp_path):
 monkeypatch.setattr(e,'COM_AVAILABLE',False)
 with pytest.raises(ValueError):e.export({'type':'calc'},str(tmp_path/'out'))

@pytest.mark.parametrize('preset',['xlsx','pptx','txt'])
def test_unrelated_routes_unchanged(monkeypatch,tmp_path,preset):
 monkeypatch.setattr(e,'COM_AVAILABLE',False)
 with pytest.raises(RuntimeError,match='WPS COM'):e.export({},str(tmp_path/'out'),preset)

@pytest.mark.parametrize('preset',['docx','pdf'])
def test_com_path(monkeypatch,tmp_path,preset):
 monkeypatch.setattr(e,'COM_AVAILABLE',True)
 app=Mock();doc=Mock();calls=[]
 monkeypatch.setattr(e,'find_wps',lambda kind:app,raising=False)
 monkeypatch.setattr(e,'create_document',lambda a,kind:doc,raising=False)
 monkeypatch.setattr(e,'_fill_document',lambda *args:calls.append('fill'))
 def save(d,path,*args):
  from pathlib import Path
  Path(path).write_bytes(b'COM mock');calls.append('save')
 for name,value in [('save_as',save),('export_pdf',save),('close_document',Mock()),('quit_app',Mock()),('get_version',lambda a:'mock')]:monkeypatch.setattr(e,name,value,raising=False)
 result=e.export({'type':'writer'},str(tmp_path/'out'),preset)
 assert result['method']=='wps-com-automation';assert calls==['fill','save']
