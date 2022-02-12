from setuptools import setup

APP = ['1st_order.py']
APP_NAME = "1st Order"
DATA_FILES = []
OPTIONS = {'argv_emulation': True,
           #'iconfile': 'logo.png',
           'plist': {
               'CFBundleName': APP_NAME,
               'CFBundleDisplayName': APP_NAME,
               'CFBundleGetInfoString': "1st Orders",
               'CFBundleIdentifier': "com.palash.osx.e-mutation",
               'CFBundleVersion': "0.1.0",
               'CFBundleShortVersionString': "0.1.0",
               'NSHumanReadableCopyright': u"Copyright © 2020, Palash Mondal, All Rights Reserved"
           }
        }
setup(
    name = APP_NAME,
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)