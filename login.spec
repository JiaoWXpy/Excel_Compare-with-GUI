# -*- mode: python -*-

block_cipher = None


a = Analysis(['login.py'],
             pathex=['homepage.py', 'C:\\Users\\Administrator\\Desktop\\Gifolye_exe'],
             binaries=[],
             datas=[],
             hiddenimports=['homepage'],
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
          [],
          name='login',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='bp.ico')
