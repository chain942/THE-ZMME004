*&---------------------------------------------------------------------*
*& Report YZMME004_TESTDATA
*&---------------------------------------------------------------------*
* READ ONLY - SELECT เท่านั้น ไม่มี UPDATE/INSERT/DELETE
* สร้างใน $TMP ไม่ต้อง transport
* ผลลัพธ์: D:\for_query\zmme004.json
*----------------------------------------------------------------------*
REPORT yzmme004_testdata.

DATA gt_out TYPE string_table.

DEFINE dump.
  APPEND |### { &1 }| TO gt_out.
  APPEND /ui2/cl_json=>serialize( &2 ) TO gt_out.
END-OF-DEFINITION.

START-OF-SELECTION.

  " batch level / ระดับ plant หรือ material
  SELECT * FROM tcuch INTO TABLE @DATA(lt_tcuch).
  dump 'TCUCH batch level' lt_tcuch.

  " batch class (023)
  SELECT clint, class, statu
    FROM klah
    WHERE klart = '023'
    INTO TABLE @DATA(lt_klah).
  dump 'KLAH class type 023' lt_klah.

  SELECT klah~class, cabn~atnam, cabn~atfor, cabn~atein, cabn~anzst, cabnt~atbez, cabnt~spras
    FROM klah
    INNER JOIN ksml ON ksml~clint = klah~clint
    INNER JOIN cabn ON cabn~atinn = ksml~imerk
    LEFT OUTER JOIN cabnt ON  cabnt~atinn = cabn~atinn
                          AND cabnt~adzhl = cabn~adzhl
    WHERE klah~klart = '023'
    ORDER BY klah~class, cabn~atnam
    INTO TABLE @DATA(lt_chars).
  dump 'Characteristics of 023 classes' lt_chars.

  " material -> batch class
  SELECT inob~objek, inob~obtab, inob~cuobj
    FROM inob
    WHERE klart = '023'
      AND obtab = 'MARA'
    INTO TABLE @DATA(lt_inob_mara)
    UP TO 30 ROWS.
  dump 'INOB MARA 023' lt_inob_mara.

  SELECT COUNT(*) FROM inob WHERE klart = '023' AND obtab = 'MCH1' INTO @DATA(lv_mch1).
  SELECT COUNT(*) FROM inob WHERE klart = '023' AND obtab = 'MCHA' INTO @DATA(lv_mcha).
  dump 'INOB count MCH1' lv_mch1.
  dump 'INOB count MCHA' lv_mcha.

  SELECT objek, mafid, klart, clint
    FROM kssk
    WHERE klart = '023'
    INTO TABLE @DATA(lt_kssk)
    UP TO 30 ROWS.
  dump 'KSSK 023' lt_kssk.

  " batch ตัวอย่าง
  SELECT mcha~matnr, mcha~werks, mcha~charg, mara~mtart
    FROM mcha
    INNER JOIN mara ON mara~matnr = mcha~matnr
    INTO TABLE @DATA(lt_batch)
    UP TO 30 ROWS.
  dump 'MCHA sample' lt_batch.

  SELECT mtart, COUNT(*) AS cnt
    FROM mara
    WHERE mtart IN ( 'Z10', 'Z20', 'Z50' )
    GROUP BY mtart
    INTO TABLE @DATA(lt_mtart).
  dump 'MARA count by type Z10/Z20/Z50' lt_mtart.

  " ของเดิม GJS มีบนระบบนี้ไหม
  SELECT funcname, pname
    FROM tfdir
    WHERE funcname IN ( 'ZCA_UPDATE_BATCH_CHAR', 'ZCA_GET_OBJ_CLASS', 'ZCA_GET_BATCH_CLASS' )
    INTO TABLE @DATA(lt_fm).
  dump 'TFDIR ZCA FMs' lt_fm.

  SELECT objct, oclss
    FROM tobj
    WHERE objct IN ( 'ZMME03_SD', 'ZMME03_MM', 'ZMME03_QM', 'M_MATE_WRK' )
    INTO TABLE @DATA(lt_tobj).
  dump 'TOBJ authorization objects' lt_tobj.

  SELECT werks, name1 FROM t001w INTO TABLE @DATA(lt_plant).
  dump 'T001W plants' lt_plant.

  cl_gui_frontend_services=>gui_download(
    EXPORTING
      filename = 'D:\for_query\zmme004.json'
      codepage = '4110'
    CHANGING
      data_tab = gt_out
    EXCEPTIONS
      OTHERS   = 1 ).

  IF sy-subrc = 0.
    MESSAGE 'D:\for_query\zmme004.json written' TYPE 'S'.
  ELSE.
    MESSAGE 'Download failed' TYPE 'S' DISPLAY LIKE 'E'.
  ENDIF.
