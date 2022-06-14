
@Echo off

@REM SET mypath="\\SRV-DC01.sve.local\Sve_Datas\INFORMATIQUE\IT\Mise en Page Excel"

@REM pushd %mypath%
echo Lancement de l'application, merci de patienter...
@REM rem cd /d %mypath%
CALL .\notebookenv\Scripts\activate.bat
python "main.py"
@REM popd
