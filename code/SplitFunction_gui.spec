# -*- mode: python -*-

block_cipher = None

gooey_languages = Tree('C:/Users/1340041/AppData/Local/Continuum/anaconda3/Lib/site-packages/gooey/languages', prefix = 'gooey/languages')
gooey_images = Tree('C:/Users/1340041/AppData/Local/Continuum/anaconda3/Lib/site-packages/gooey/images', prefix = 'gooey/images')

a = Analysis(['SplitFunction_gui.py'],
             pathex=['D:\\Documents and Settings\\1340041\\Desktop\\Program'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          gooey_images,
          gooey_languages,
          [],
          name='DispatchAtt',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
