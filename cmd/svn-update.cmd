set vCmdHome=C:\Program Files\TortoiseSVN\bin\
"%vCmdHome%"TortoiseProc /command:update /path:%1

@rem // set command home
set vCmdHome=%userprofile%\Documents\tools\cmd

@rem // file list all to csv
start %vCmdHome%\cre-csv-filepath.cmd "C:\Users\winridge\Documents\workspace\svn-work\ドキュメント\再開発" "file_lists_all.csv"
@rem // file list spec to csv
start %vCmdHome%\cre-csv-filepath.cmd "C:\Users\winridge\Documents\workspace\svn-work\ドキュメント\再開発\phase_201411\15_受領資料" "file_lists_doc.csv"
