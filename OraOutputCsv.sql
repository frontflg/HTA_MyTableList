CREATE OR REPLACE PROCEDURE OutputCsv (vTableName IN VARCHAR2)
IS
-- vColumns    VARCHAR2(32767);
   vSelectSql  VARCHAR2(32767);
   vFirstFlag  BOOLEAN       := TRUE; -- 先頭区切り文字スキップ
-- CREATE OR REPLACE DIRECTORY temp_dir AS 'C:\temp';
-- GRANT READ, WRITE ON DIRECTORY temp_dir TO user;
   vDirname    VARCHAR2(200) := 'temp_dir';
   vFilename   VARCHAR2(200) := 'TEST.CSV';
   vHandle     UTL_FILE.FILE_TYPE;

   TYPE        cusorType IS REF CURSOR;
   c           cusorType;
   vOutputLine VARCHAR2(32767);

BEGIN
-- vColumns   := NULL;
   vSelectSql := 'SELECT ';
-- vSelectSql := 'SELECT ''"''';
   FOR vRec IN (
      SELECT COLUMN_NAME
      FROM USER_TAB_COLUMNS
      WHERE TABLE_NAME = UPPER(vTableName)
      ORDER BY COLUMN_ID
   ) LOOP
      IF NOT(vFirstFlag) THEN
      -- vColumns   := vColumns || ',';
         vSelectSql := vSelectSql || ' || '','' || ';
      -- vSelectSql := vSelectSql || ' || ''","'' || ';
      ELSE
         vFirstFlag := FALSE;
      END IF;
   -- vColumns      := vColumns   || vRec.COLUMN_NAME;
      vSelectSql    := vSelectSql || vRec.COLUMN_NAME;
   -- vSelectSql    := vSelectSql || '"' || vRec.COLUMN_NAME || '"';
   -- CONCAT('"',item1,'","','"',item2,'"') item12 FROM tableid;
   END LOOP;
   vSelectSql := vSelectSql || ' FROM ' || vTableName;
   vFilename  := vTableName || '.csv';

   vHandle := UTL_FILE.FOPEN(vDirname, vFilename, 'W', 32767);
-- DBMS_OUTPUT.PUT_LINE(vColumns);
-- UTL_FILE.PUT_LINE(vHandle,vColumns);
   OPEN c FOR vSelectSql;
   LOOP
      FETCH c INTO vOutputLine;
      EXIT WHEN c%NOTFOUND;
   -- DBMS_OUTPUT.PUT_LINE(vOutputLine);
      UTL_FILE.PUT_LINE(vHandle,vOutputLine);
   END LOOP;
   CLOSE c;
   UTL_FILE.FCLOSE(vHandle);

EXCEPTION WHEN OTHERS THEN
   UTL_FILE.FCLOSE_ALL;
   RAISE;
END;
/
SHOW ERROR

■csvout.bat
@echo off
sqlplus -s user/password@dbname @csvout.sql

■csvout.sql
SET serveroutput ON;
BEGIN
   OWNER.OraOutputCsv("○○マスタ");
   OWNER.OraOutputCsv("◇◇テーブル");
END;
/
exit;
