import pyrpm

# Создаем объект RPM пакета
pkg = pyrpm.RPM('ogti_log_pas')

# Set the package version
pkg.version = '1.0'

# Set the package summary and description
pkg.summary = 'OGTI Log Pas package'
pkg.description = 'This is a package for OGTI Log Pas'

# Add the main.py file to the package
pkg.add_file('C:\\Users\\Prince_GG\\Desktop\\OGTI_Log_Pas\\main.py', '/usr/lib/python3.8/site-packages/ogti_log_pas/')

# Create the RPM package
pkg.build()
