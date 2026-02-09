from cx_Freeze import setup, Executable

setup(
    name="ArrumarExtratosV1",
    version="1.0",
    description="Your application description",
    executables=[Executable("app.py", base=None)]
)