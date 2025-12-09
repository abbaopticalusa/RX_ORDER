# app.spec (수정된 최종 버전)
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['app.py'],
             pathex=['.'],
             binaries=[],
             # Base64 파일을 EXE에 데이터로 포함시킵니다.
             datas=[('excel_template.txt', '.')], 
             # 🚨 숨겨진 종속성 추가: openpyxl, pandas, streamlit 관련 오류 방지
             hiddenimports=['openpyxl.worksheet._read_only', 'openpyxl.xml.constants', 'pandas._libs.tslibs.timedeltas', 'streamlit'], 
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='PlazmaOrderApp',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False, # 실행 시 검은색 콘솔 창 숨기기
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )