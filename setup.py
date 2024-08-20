from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': ['os','xlwings','datetime'], 'excludes': []}

base = 'console'

executables = [
    Executable('main.py', base=base)
]

setup(name='Disney Offers Tool',
      version = '1.0',
      description = 'A one click tool for Disney Offers',
      options = {'build_exe': build_options},
      executables = executables)
