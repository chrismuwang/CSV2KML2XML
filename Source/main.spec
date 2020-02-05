# -*- mode: python -*-

block_cipher = None


a = Analysis(['CSV2KML2XML.py'],
             pathex=['C:\\Python36\\Scripts'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

a.datas += [('Instructions.pdf','Instructions.pdf', "DATA"), ('S&C Icon.ico','S&C Icon.ico', "DATA")]

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='CSV2KML2XML v5',
          debug=False,
          strip=False,
          upx=True,
          console=False,
          icon='S&C Icon.ico')