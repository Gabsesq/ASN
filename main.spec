# main.spec
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates', 'templates'),  # Add entire templates directory
        ('static', 'static'),
        ('assets/Chewy', 'assets/Chewy'),
        ('assets/Chewy/Chewy 856 ASN - Copy.xlsx', 'assets/Chewy'),
        ('assets/Chewy/Chewy UCC128 Label Request - Copy.xlsx', 'assets/Chewy'),

        ('assets/Murdochs', 'assets/Murdochs'),

        ('assets/Pet Supermarket', 'assets/Pet Supermarket'),

        ('assets/Scheels', 'assets/Scheels'),
        ('assets/Scheels/Blank Scheels 856 ASN.xlsx', 'assets/Scheels'),
        ('assets/Scheels/Blank Scheels UCC128 Label Request.xlsx', 'assets/Scheels'),

        ('assets/Thrive Market', 'assets/Thrive Market'),
        ('assets/Thrive Market/Blank Thrive Market UCC128 Label Request 7.19.24.xlsx', 'assets/Thrive Market'),
        ('assets/Thrive Market/Thrive Market 856 Master Template.xlsx', 'assets/Thrive Market'),

        ('assets/TSC', 'assets/TSC'),
        ('assets/TSC/Blank TSC ASN.xlsx', 'assets/TSC'),

        ('assets/TSC IS', 'assets/TSC IS'),

        ('Finished/Chewy', 'Finished/Chewy'),
        ('Finished/Murdochs', 'Finished/Murdochs'),
        ('Finished/PetSupermarket', 'Finished/PetSupermarket'),
        ('Finished/Thrive', 'Finished/Thrive'),
        ('Finished/TSC', 'Finished/TSC')
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)