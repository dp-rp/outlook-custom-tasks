import os, sysconfig
print(sysconfig.get_path('scripts',f'{os.name}_user'))
