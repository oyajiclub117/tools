set vCmdHome=C:\Program Files\TortoiseSVN\bin\
"%vCmdHome%"TortoiseProc /command:update /path:%1

@rem // set command home
set vCmdHome=%userprofile%\Documents\tools\cmd

@rem // file list all to csv
start %vCmdHome%\cre-csv-filepath.cmd "C:\Users\winridge\Documents\workspace\svn-work\�h�L�������g\�ĊJ��" "file_lists_all.csv"
@rem // file list spec to csv
start %vCmdHome%\cre-csv-filepath.cmd "C:\Users\winridge\Documents\workspace\svn-work\�h�L�������g\�ĊJ��\phase_201411\15_��̎���" "file_lists_doc.csv"
