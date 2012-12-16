from cx_Freeze import setup, Executable

exe = Executable(
    script="Helper.pyw",
    icon = "icons/icon.ico",
    base="Win32GUI")

#build_options = {"packages": ["xlrd"]}

setup(
        name = "Helper",
        version = "0.1",
        description = "Dictionary raw data formatter helper",
        #options = {"build_exe": build_options},
        executables = [exe])