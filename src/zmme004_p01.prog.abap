*&---------------------------------------------------------------------*
*& Include          ZMME004_P01
*&---------------------------------------------------------------------*
CLASS lcx_error IMPLEMENTATION.

  METHOD constructor.
    super->constructor( ).
    mv_text = iv_text.
  ENDMETHOD.

  METHOD if_message~get_text.
    result = mv_text.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Excel workbook (.xlsx) from the front end
*&---------------------------------------------------------------------*
CLASS lcl_workbook IMPLEMENTATION.

  METHOD is_open.
    rv_open = xsdbool( go_excel IS BOUND AND gv_file = iv_file ).
  ENDMETHOD.

  METHOD sheet_names.
    open( iv_file )->if_fdt_doc_spreadsheet~get_worksheet_names(
      IMPORTING worksheet_names = DATA(lt_names) ).

    rt_names = VALUE #( FOR lv_name IN lt_names ( CONV #( lv_name ) ) ).
  ENDMETHOD.

  METHOD read_sheet.
    FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.

    DATA(lo_excel) = open( iv_file ).

    DATA(lv_sheet) = iv_sheet.
    IF lv_sheet IS INITIAL.
      DATA(lt_names) = sheet_names( iv_file ).
      lv_sheet = VALUE #( lt_names[ 1 ] OPTIONAL ).
    ENDIF.

    DATA(lr_sheet) = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lv_sheet ).
    IF lr_sheet IS NOT BOUND.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING iv_text = |Worksheet { lv_sheet } not found|.
    ENDIF.

    ASSIGN lr_sheet->* TO <lt_sheet>.

    LOOP AT <lt_sheet> ASSIGNING FIELD-SYMBOL(<ls_line>).
      APPEND INITIAL LINE TO rt_rows ASSIGNING FIELD-SYMBOL(<lt_cells>).

      DO.
        ASSIGN COMPONENT sy-index OF STRUCTURE <ls_line> TO FIELD-SYMBOL(<lv_cell>).
        IF sy-subrc <> 0.
          EXIT.
        ENDIF.

        DATA(lv_cell) = condense( val  = CONV string( <lv_cell> )
                                  from = `` ).
        APPEND lv_cell TO <lt_cells>.
      ENDDO.
    ENDLOOP.
  ENDMETHOD.

  METHOD open.
    IF is_open( iv_file ).
      ro_excel = go_excel.
      RETURN.
    ENDIF.

    DATA(lv_xstring) = upload( iv_file ).

    TRY.
        go_excel = NEW cl_fdt_xl_spreadsheet( document_name = iv_file
                                              xdocument     = lv_xstring ).
      CATCH cx_fdt_excel_core.
        CLEAR: go_excel, gv_file.
        RAISE EXCEPTION TYPE lcx_error
          EXPORTING iv_text = |{ iv_file } is not an Excel workbook (.xlsx)|.
    ENDTRY.

    gv_file  = iv_file.
    ro_excel = go_excel.
  ENDMETHOD.

  METHOD upload.
    DATA: lt_binary TYPE solix_tab,
          lv_size   TYPE i.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename   = iv_file
        filetype   = 'BIN'
      IMPORTING
        filelength = lv_size
      CHANGING
        data_tab   = lt_binary
      EXCEPTIONS
        OTHERS     = 1 ).

    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING iv_text = |Cannot read file { iv_file }|.
    ENDIF.

    rv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_binary
                                                   iv_size  = lv_size ).
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Batch classification (class type 023)
*&---------------------------------------------------------------------*
CLASS lcl_batch_char IMPLEMENTATION.

  METHOD batch_class.
    DATA: lv_objek      TYPE inob-objek,
          lv_kssk_objek TYPE kssk-objek.

    READ TABLE mt_class_cache INTO DATA(ls_cache) WITH TABLE KEY matnr = iv_matnr.
    IF sy-subrc = 0.
      rv_class = ls_cache-class.
      RETURN.
    ENDIF.

    lv_objek = iv_matnr.

    SELECT SINGLE cuobj
      FROM inob
      WHERE klart = @c_klart
        AND obtab = 'MARA'
        AND objek = @lv_objek
      INTO @DATA(lv_cuobj).

    IF sy-subrc = 0.
      lv_kssk_objek = lv_cuobj.

      SELECT SINGLE klah~class
        FROM kssk
        INNER JOIN klah ON klah~clint = kssk~clint
        WHERE kssk~objek = @lv_kssk_objek
          AND kssk~mafid = 'O'
          AND kssk~klart = @c_klart
        INTO @rv_class.
    ENDIF.

    INSERT VALUE #( matnr = iv_matnr
                    class = rv_class ) INTO TABLE mt_class_cache.
  ENDMETHOD.

  METHOD characteristic_list.
    DATA lt_class TYPE RANGE OF klasse_d.

    IF iv_class IS NOT INITIAL.
      lt_class = VALUE #( ( sign = 'I' option = 'EQ' low = iv_class ) ).
    ENDIF.

    SELECT DISTINCT cabn~atnam, cabnt~atbez
      FROM klah
      INNER JOIN ksml ON ksml~clint = klah~clint
      INNER JOIN cabn ON cabn~atinn = ksml~imerk
      LEFT OUTER JOIN cabnt ON  cabnt~atinn = cabn~atinn
                            AND cabnt~adzhl = cabn~adzhl
                            AND cabnt~spras = @c_text_lang
      WHERE klah~klart = @c_klart
        AND klah~class IN @lt_class
      ORDER BY cabn~atnam
      INTO TABLE @rt_list.
  ENDMETHOD.

  METHOD descriptions.
    DATA lt_atnam TYPE RANGE OF atnam.

    IF it_chars IS INITIAL.
      RETURN.
    ENDIF.

    lt_atnam = VALUE #( FOR lv_char IN it_chars
                        ( sign = 'I' option = 'EQ' low = lv_char ) ).

    SELECT cabn~atnam, cabnt~atbez
      FROM cabn
      INNER JOIN cabnt ON  cabnt~atinn = cabn~atinn
                       AND cabnt~adzhl = cabn~adzhl
      WHERE cabn~atnam  IN @lt_atnam
        AND cabnt~spras =  @c_text_lang
      INTO TABLE @rt_text.
  ENDMETHOD.

  METHOD class_chars.
    READ TABLE mt_char_cache INTO DATA(ls_cache) WITH TABLE KEY class = iv_class.
    IF sy-subrc = 0.
      rt_chars = ls_cache-chars.
      RETURN.
    ENDIF.

    SELECT DISTINCT cabn~atnam, cabn~atfor, cabn~atein
      FROM klah
      INNER JOIN ksml ON ksml~clint = klah~clint
      INNER JOIN cabn ON cabn~atinn = ksml~imerk
      WHERE klah~klart = @c_klart
        AND klah~class = @iv_class
      INTO TABLE @rt_chars.

    INSERT VALUE #( class = iv_class
                    chars = rt_chars ) INTO TABLE mt_char_cache.
  ENDMETHOD.

  METHOD update.
    DATA: lv_key    TYPE bapi1003_key-object_long,
          ls_values TYPE ty_values,
          lt_return TYPE bapiret2_t.

    DATA(lv_class) = batch_class( iv_matnr ).
    IF lv_class IS INITIAL.
      rs_result-status = |Material { iv_matnr ALPHA = OUT } has no batch class (023)|.
      RETURN.
    ENDIF.

    DATA(lt_input) = split_values( it_values ).
    DELETE lt_input WHERE value IS INITIAL.
    IF lt_input IS INITIAL.
      rs_result-status = 'No value to update'.
      RETURN.
    ENDIF.

    lv_key    = iv_matnr.
    lv_key+40 = iv_charg.

    " ค่าปัจจุบันของ batch
    CALL FUNCTION 'BAPI_OBJCL_GETDETAIL'
      EXPORTING
        objecttable     = 'MCH1'
        classnum        = lv_class
        classtype       = c_klart
        objectkey_long  = lv_key
      TABLES
        allocvaluesnum  = ls_values-num
        allocvalueschar = ls_values-char
        allocvaluescurr = ls_values-curr
        return          = lt_return.

    " ใส่ค่าใหม่
    DATA(lt_chars) = class_chars( lv_class ).

    LOOP AT lt_input INTO DATA(ls_input).
      READ TABLE lt_chars INTO DATA(ls_char) WITH TABLE KEY atnam = ls_input-charact.
      IF sy-subrc <> 0.
        rs_result-status = |{ ls_input-charact } is not in class { lv_class }|.
        RETURN.
      ENDIF.

      TRY.
          apply_value( EXPORTING is_char   = ls_char
                                 iv_value  = ls_input-value
                       CHANGING  cs_values = ls_values ).
        CATCH lcx_error INTO DATA(lx_error).
          rs_result-status = lx_error->get_text( ).
          RETURN.
      ENDTRY.
    ENDLOOP.

    " บันทึก
    CLEAR lt_return.

    CALL FUNCTION 'BAPI_OBJCL_CHANGE'
      EXPORTING
        objecttable        = 'MCH1'
        classnum           = lv_class
        classtype          = c_klart
        objectkey_long     = lv_key
      TABLES
        allocvaluesnumnew  = ls_values-num
        allocvaluescharnew = ls_values-char
        allocvaluescurrnew = ls_values-curr
        return             = lt_return.

    LOOP AT lt_return INTO DATA(ls_return) WHERE type = 'E'
                                              OR type = 'A'.
      rs_result-status = ls_return-message.
      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
      RETURN.
    ENDLOOP.

    CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
      EXPORTING
        wait = abap_true.

    rs_result = VALUE #( ok     = abap_true
                         status = 'Update batch classification completed' ).
  ENDMETHOD.

  METHOD split_values.
    LOOP AT it_values INTO DATA(ls_value).
      IF ls_value-value NS c_separator.
        APPEND ls_value TO rt_values.
        CONTINUE.
      ENDIF.

      SPLIT ls_value-value AT c_separator INTO TABLE DATA(lt_parts).

      LOOP AT lt_parts INTO DATA(lv_part).
        APPEND VALUE #( charact = ls_value-charact
                        value   = condense( lv_part ) ) TO rt_values.
      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.

  METHOD apply_value.
    IF to_upper( iv_value ) = c_blank.
      DELETE cs_values-char WHERE charact = is_char-atnam.
      DELETE cs_values-num  WHERE charact = is_char-atnam.
      DELETE cs_values-curr WHERE charact = is_char-atnam.
      RETURN.
    ENDIF.

    CASE is_char-atfor.
      WHEN 'CHAR' OR space.
        set_char_value( EXPORTING is_char  = is_char
                                  iv_value = iv_value
                        CHANGING  ct_char  = cs_values-char ).

      WHEN 'CURR'.
        DATA(lv_amount) = to_number( is_char  = is_char
                                     iv_value = iv_value ).

        READ TABLE cs_values-curr ASSIGNING FIELD-SYMBOL(<ls_curr>)
          WITH KEY charact = is_char-atnam.
        IF sy-subrc <> 0 OR is_char-atein IS INITIAL.
          APPEND VALUE #( charact = is_char-atnam ) TO cs_values-curr ASSIGNING <ls_curr>.
        ENDIF.

        <ls_curr>-value_from = lv_amount.

      WHEN OTHERS.
        DATA(lv_number) = to_number( is_char  = is_char
                                     iv_value = iv_value ).

        READ TABLE cs_values-num ASSIGNING FIELD-SYMBOL(<ls_num>)
          WITH KEY charact = is_char-atnam.
        IF sy-subrc <> 0 OR is_char-atein IS INITIAL.
          APPEND VALUE #( charact = is_char-atnam ) TO cs_values-num ASSIGNING <ls_num>.
        ENDIF.

        <ls_num>-value_from = lv_number.
    ENDCASE.
  ENDMETHOD.

  METHOD set_char_value.
    READ TABLE ct_char ASSIGNING FIELD-SYMBOL(<ls_char>)
      WITH KEY charact = is_char-atnam.

    " single value ทับของเดิม, multiple value เพิ่มต่อท้าย
    IF sy-subrc <> 0 OR is_char-atein IS INITIAL.
      IF line_exists( ct_char[ charact            = is_char-atnam
                               value_neutral_long = iv_value ] ).
        RETURN.
      ENDIF.

      APPEND VALUE #( charact = is_char-atnam ) TO ct_char ASSIGNING <ls_char>.
    ENDIF.

    <ls_char>-value_char         = iv_value.
    <ls_char>-value_neutral      = iv_value.
    <ls_char>-value_char_long    = iv_value.
    <ls_char>-value_neutral_long = iv_value.
  ENDMETHOD.

  METHOD to_number.
    DATA(lv_text) = condense( iv_value ).

    IF is_char-atfor = 'DATE'.
      lv_text = to_date_text( lv_text ).
    ENDIF.

    TRY.
        rv_value = lv_text.
      CATCH cx_sy_conversion_error.
        RAISE EXCEPTION TYPE lcx_error
          EXPORTING iv_text = |{ is_char-atnam }: { iv_value } is not a valid value|.
    ENDTRY.
  ENDMETHOD.

  METHOD to_date_text.
    DATA lv_date TYPE d.

    " Excel ส่งวันที่มาเป็นเลข serial
    IF matches( val = iv_text regex = `\d{1,5}` ).
      lv_date = CONV d( '18991230' ) + CONV i( iv_text ).
      rv_text = lv_date.

    ELSEIF matches( val = iv_text regex = `\d{2}\.\d{2}\.\d{4}` ).
      rv_text = substring( val = iv_text off = 6 len = 4 )
             && substring( val = iv_text off = 3 len = 2 )
             && substring( val = iv_text off = 0 len = 2 ).

    ELSE.
      rv_text = iv_text.
    ENDIF.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Selection screen events
*&---------------------------------------------------------------------*
CLASS lcl_screen IMPLEMENTATION.

  METHOD init.
    b_clr1 = b_clr2 = b_clr3 = b_clr4 = b_clr5 = icon_cancel.
  ENDMETHOD.

  METHOD pbo.
    LOOP AT SCREEN INTO DATA(ls_screen).
      IF ( ls_screen-group1 = 'SCR' AND r_xls = abap_true )
      OR ( ls_screen-group1 = 'XLS' AND r_scr = abap_true ).
        ls_screen-active = 0.
        MODIFY SCREEN FROM ls_screen.
      ENDIF.
    ENDLOOP.

    IF r_xls = abap_true AND lcl_workbook=>is_open( CONV #( p_file ) ).
      fill_sheet_list( ).
    ENDIF.
  ENDMETHOD.

  METHOD pai.
    CASE sscrfields-ucomm.
      WHEN 'CLR1'.
        CLEAR: p_atn1, p_val1.
      WHEN 'CLR2'.
        CLEAR: p_atn2, p_val2.
      WHEN 'CLR3'.
        CLEAR: p_atn3, p_val3.
      WHEN 'CLR4'.
        CLEAR: p_atn4, p_val4.
      WHEN 'CLR5'.
        CLEAR: p_atn5, p_val5.
      WHEN 'ONLI'.
        check_input( ).
    ENDCASE.
  ENDMETHOD.

  METHOD f4_file.
    DATA: lt_files TYPE filetable,
          lv_rc    TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        window_title = 'Select an Excel workbook'
        file_filter  = 'Excel Workbook (*.xlsx)|*.xlsx'
      CHANGING
        file_table   = lt_files
        rc           = lv_rc
      EXCEPTIONS
        OTHERS       = 1 ).

    IF sy-subrc <> 0 OR lv_rc < 1.
      RETURN.
    ENDIF.

    p_file = lt_files[ 1 ]-filename.

    update_screen_field( iv_field = 'P_SHEET'
                         iv_value = fill_sheet_list( ) ).
  ENDMETHOD.

  METHOD f4_characteristic.
    DATA lv_class TYPE klasse_d.

    DATA(lo_classif) = NEW lcl_batch_char( ).

    " กรองตาม batch class ของ material
    DATA(lv_matnr) = lcl_app=>to_matnr( read_screen_field( 'P_MATNR' ) ).
    IF lv_matnr IS NOT INITIAL.
      lv_class = lo_classif->batch_class( lv_matnr ).
    ENDIF.

    DATA(lt_list)  = lo_classif->characteristic_list( lv_class ).
    DATA(lv_repid) = sy-repid.

    CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST'
      EXPORTING
        retfield    = 'ATNAM'
        dynpprog    = lv_repid
        dynpnr      = sy-dynnr
        dynprofield = iv_field
        value_org   = 'S'
      TABLES
        value_tab   = lt_list
      EXCEPTIONS
        OTHERS      = 1.
  ENDMETHOD.

  METHOD char_values.
    rt_values = VALUE #( ( charact = p_atn1  value = p_val1 )
                         ( charact = p_atn2  value = p_val2 )
                         ( charact = p_atn3  value = p_val3 )
                         ( charact = p_atn4  value = p_val4 )
                         ( charact = p_atn5  value = p_val5 ) ).

    DELETE rt_values WHERE charact IS INITIAL.
  ENDMETHOD.

  METHOD check_input.
    DATA lv_message TYPE string.

    IF r_xls = abap_true.
      IF p_file IS INITIAL.
        MESSAGE 'Enter the file name' TYPE 'E'.
      ENDIF.
      RETURN.
    ENDIF.

    IF p_matnr IS INITIAL
    OR p_werks IS INITIAL
    OR s_charg[] IS INITIAL.
      MESSAGE 'Enter material, plant and batch' TYPE 'E'.
    ENDIF.

    DATA(lt_values) = char_values( ).
    IF lt_values IS INITIAL.
      MESSAGE 'Enter at least one characteristic' TYPE 'E'.
    ENDIF.

    LOOP AT lt_values INTO DATA(ls_value) WHERE value IS INITIAL.
      lv_message = |Enter a value for { ls_value-charact } (BLANK clears it)|.
      MESSAGE lv_message TYPE 'E'.
    ENDLOOP.
  ENDMETHOD.

  METHOD fill_sheet_list.
    TRY.
        DATA(lt_names) = lcl_workbook=>sheet_names( CONV #( p_file ) ).
      CATCH lcx_error INTO DATA(lx_error).
        MESSAGE lx_error TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
    ENDTRY.

    DATA(lt_list) = VALUE vrm_values( FOR lv_name IN lt_names
                                      ( key = lv_name  text = lv_name ) ).

    CALL FUNCTION 'VRM_SET_VALUES'
      EXPORTING
        id     = 'P_SHEET'
        values = lt_list
      EXCEPTIONS
        OTHERS = 1.

    rv_sheet = p_sheet.
    IF NOT line_exists( lt_names[ table_line = rv_sheet ] ).
      rv_sheet = VALUE #( lt_names[ 1 ] OPTIONAL ).
    ENDIF.

    p_sheet = rv_sheet.
  ENDMETHOD.

  METHOD read_screen_field.
    DATA lt_fields TYPE STANDARD TABLE OF dynpread WITH EMPTY KEY.

    DATA(lv_repid) = sy-repid.
    lt_fields = VALUE #( ( fieldname = iv_field ) ).

    CALL FUNCTION 'DYNP_VALUES_READ'
      EXPORTING
        dyname     = lv_repid
        dynumb     = sy-dynnr
      TABLES
        dynpfields = lt_fields
      EXCEPTIONS
        OTHERS     = 1.

    IF sy-subrc = 0.
      rv_value = VALUE #( lt_fields[ 1 ]-fieldvalue OPTIONAL ).
    ENDIF.
  ENDMETHOD.

  METHOD update_screen_field.
    DATA lt_fields TYPE STANDARD TABLE OF dynpread WITH EMPTY KEY.

    DATA(lv_repid) = sy-repid.
    lt_fields = VALUE #( ( fieldname  = iv_field
                           fieldvalue = iv_value ) ).

    CALL FUNCTION 'DYNP_VALUES_UPDATE'
      EXPORTING
        dyname     = lv_repid
        dynumb     = sy-dynnr
      TABLES
        dynpfields = lt_fields
      EXCEPTIONS
        OTHERS     = 1.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Main flow: collect rows, update, show result
*&---------------------------------------------------------------------*
CLASS lcl_app IMPLEMENTATION.

  METHOD run.
    TRY.
        IF r_xls = abap_true.
          read_excel( ).
        ELSE.
          read_screen( ).
        ENDIF.
      CATCH lcx_error INTO DATA(lx_error).
        MESSAGE lx_error TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
    ENDTRY.

    IF mt_rows IS INITIAL.
      MESSAGE 'No data for updating' TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF check_plant_authority( ) = abap_false.
      RETURN.
    ENDIF.

    update_batches( ).
    show_result( ).
  ENDMETHOD.

  METHOD to_matnr.
    DATA(lv_input) = to_upper( condense( iv_text ) ).

    CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
      EXPORTING
        input        = lv_input
      IMPORTING
        output       = rv_matnr
      EXCEPTIONS
        length_error = 1
        OTHERS       = 2.

    IF sy-subrc <> 0.
      CLEAR rv_matnr.
    ENDIF.
  ENDMETHOD.

  METHOD read_excel.
    TYPES: BEGIN OF ty_column,
             index   TYPE i,
             charact TYPE atnam,
           END OF ty_column.

    DATA lt_columns TYPE STANDARD TABLE OF ty_column WITH EMPTY KEY.

    DATA(lt_sheet) = lcl_workbook=>read_sheet( iv_file  = CONV #( p_file )
                                               iv_sheet = CONV #( p_sheet ) ).

    " หัวตาราง: MATNR, PLANT, BATCH แล้วตามด้วยชื่อ characteristic
    DATA(lt_header) = VALUE string_table( lt_sheet[ 1 ] OPTIONAL ).

    LOOP AT lt_header INTO DATA(lv_header) FROM c_first_char_col.
      IF lv_header IS NOT INITIAL.
        APPEND VALUE #( index   = sy-tabix
                        charact = to_upper( lv_header ) ) TO lt_columns.
      ENDIF.
    ENDLOOP.

    IF lt_columns IS INITIAL.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING iv_text = 'Row 1 of the worksheet has no characteristic column'.
    ENDIF.

    mt_chars = VALUE #( FOR ls_col IN lt_columns ( ls_col-charact ) ).

    " ข้อมูลเริ่มแถว 2
    LOOP AT lt_sheet INTO DATA(lt_cells) FROM 2.
      DATA(lv_matnr) = VALUE string( lt_cells[ 1 ] OPTIONAL ).
      DATA(lv_plant) = VALUE string( lt_cells[ 2 ] OPTIONAL ).
      DATA(lv_batch) = VALUE string( lt_cells[ 3 ] OPTIONAL ).

      IF lv_matnr IS INITIAL AND lv_plant IS INITIAL AND lv_batch IS INITIAL.
        CONTINUE.
      ENDIF.

      APPEND INITIAL LINE TO mt_rows ASSIGNING FIELD-SYMBOL(<ls_row>).

      <ls_row>-matnr = to_matnr( lv_matnr ).
      alpha_in( EXPORTING iv_text  = to_upper( lv_plant )
                IMPORTING ev_value = <ls_row>-werks ).
      alpha_in( EXPORTING iv_text  = to_upper( lv_batch )
                IMPORTING ev_value = <ls_row>-charg ).

      <ls_row>-values = VALUE #( FOR ls_column IN lt_columns
                                 ( charact = ls_column-charact
                                   value   = VALUE #( lt_cells[ ls_column-index ] OPTIONAL ) ) ).
    ENDLOOP.
  ENDMETHOD.

  METHOD read_screen.
    DATA(lt_values) = lcl_screen=>char_values( ).

    mt_chars = VALUE #( FOR ls_value IN lt_values ( ls_value-charact ) ).

    SELECT charg
      FROM mcha
      WHERE matnr = @p_matnr
        AND werks = @p_werks
        AND charg IN @s_charg
      ORDER BY charg
      INTO TABLE @DATA(lt_batch).

    IF lt_batch IS INITIAL.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING iv_text = 'No batch found for the selection'.
    ENDIF.

    mt_rows = VALUE #( FOR ls_batch IN lt_batch
                       ( matnr  = p_matnr
                         werks  = p_werks
                         charg  = ls_batch-charg
                         values = lt_values ) ).
  ENDMETHOD.

  METHOD check_plant_authority.
    DATA lt_plant TYPE SORTED TABLE OF werks_d WITH UNIQUE KEY table_line.

    LOOP AT mt_rows INTO DATA(ls_row).
      INSERT ls_row-werks INTO TABLE lt_plant.
    ENDLOOP.

    LOOP AT lt_plant INTO DATA(lv_werks).
      AUTHORITY-CHECK OBJECT 'M_MATE_WRK'
        ID 'WERKS' FIELD lv_werks
        ID 'ACTVT' FIELD '03'.

      IF sy-subrc <> 0.
        MESSAGE s855(m3) WITH lv_werks DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
    ENDLOOP.

    rv_ok = abap_true.
  ENDMETHOD.

  METHOD update_batches.
    DATA(lo_classif) = NEW lcl_batch_char( ).
    DATA(lv_total)   = lines( mt_rows ).

    LOOP AT mt_rows ASSIGNING FIELD-SYMBOL(<ls_row>).
      cl_progress_indicator=>progress_indicate(
        i_text      = |Batch { <ls_row>-charg }|
        i_processed = sy-tabix
        i_total     = lv_total ).

      DATA(ls_result) = lo_classif->update( iv_matnr  = <ls_row>-matnr
                                            iv_charg  = <ls_row>-charg
                                            it_values = <ls_row>-values ).

      <ls_row>-icon   = COND #( WHEN ls_result-ok = abap_true
                                THEN icon_green_light
                                ELSE icon_red_light ).
      <ls_row>-status = ls_result-status.
    ENDLOOP.
  ENDMETHOD.

  METHOD show_result.
    FIELD-SYMBOLS <lt_output> TYPE STANDARD TABLE.

    DATA(lr_output) = build_output_table( ).
    ASSIGN lr_output->* TO <lt_output>.

    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lo_alv)
                                CHANGING  t_table      = <lt_output> ).

        lo_alv->get_functions( )->set_all( ).
        lo_alv->get_display_settings( )->set_striped_pattern( abap_true ).
        set_column_texts( lo_alv ).

        lo_alv->display( ).

      CATCH cx_salv_msg cx_salv_not_found INTO DATA(lx_salv).
        MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.
  ENDMETHOD.

  METHOD build_output_table.
    FIELD-SYMBOLS <lt_output> TYPE STANDARD TABLE.

    DATA(lo_string) = cl_abap_elemdescr=>get_string( ).

    " คอลัมน์คงที่ + หนึ่งคอลัมน์ต่อ characteristic
    DATA(lt_comp) = VALUE cl_abap_structdescr=>component_table(
      ( name = 'ICON'  type = CAST #( cl_abap_typedescr=>describe_by_name( 'ICON_D' ) ) )
      ( name = 'MATNR' type = CAST #( cl_abap_typedescr=>describe_by_name( 'MATNR' ) ) )
      ( name = 'WERKS' type = CAST #( cl_abap_typedescr=>describe_by_name( 'WERKS_D' ) ) )
      ( name = 'CHARG' type = CAST #( cl_abap_typedescr=>describe_by_name( 'CHARG_D' ) ) ) ).

    DO lines( mt_chars ) TIMES.
      APPEND VALUE #( name = column_name( sy-index )
                      type = lo_string ) TO lt_comp.
    ENDDO.

    APPEND VALUE #( name = 'STATUS'
                    type = lo_string ) TO lt_comp.

    DATA(lo_table) = cl_abap_tabledescr=>create( cl_abap_structdescr=>get( lt_comp ) ).
    CREATE DATA rr_table TYPE HANDLE lo_table.
    ASSIGN rr_table->* TO <lt_output>.

    LOOP AT mt_rows INTO DATA(ls_row).
      APPEND INITIAL LINE TO <lt_output> ASSIGNING FIELD-SYMBOL(<ls_output>).
      MOVE-CORRESPONDING ls_row TO <ls_output>.

      LOOP AT mt_chars INTO DATA(lv_char).
        DATA(lv_column) = column_name( sy-tabix ).
        ASSIGN COMPONENT lv_column OF STRUCTURE <ls_output> TO FIELD-SYMBOL(<lv_value>).
        <lv_value> = VALUE string( ls_row-values[ charact = lv_char ]-value OPTIONAL ).
      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.

  METHOD set_column_texts.
    DATA(lo_columns) = io_alv->get_columns( ).
    lo_columns->set_optimize( ).

    DATA(lo_icon) = CAST cl_salv_column_table( lo_columns->get_column( 'ICON' ) ).
    lo_icon->set_icon( abap_true ).

    set_column_text( io_column = lo_columns->get_column( 'ICON' )
                     iv_text   = 'Type' ).
    set_column_text( io_column = lo_columns->get_column( 'STATUS' )
                     iv_text   = 'Message' ).

    DATA(lt_text) = NEW lcl_batch_char( )->descriptions( mt_chars ).

    LOOP AT mt_chars INTO DATA(lv_char).
      DATA(lv_column) = column_name( sy-tabix ).
      DATA(lv_text)   = VALUE atbez( lt_text[ atnam = lv_char ]-atbez DEFAULT lv_char ).

      set_column_text( io_column = lo_columns->get_column( lv_column )
                       iv_text   = lv_text ).
    ENDLOOP.
  ENDMETHOD.

  METHOD set_column_text.
    io_column->set_short_text( CONV #( iv_text ) ).
    io_column->set_medium_text( CONV #( iv_text ) ).
    io_column->set_long_text( CONV #( iv_text ) ).
  ENDMETHOD.

  METHOD column_name.
    rv_name = |C{ iv_index WIDTH = 3 ALIGN = RIGHT PAD = '0' }|.
  ENDMETHOD.

  METHOD alpha_in.
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = iv_text
      IMPORTING
        output = ev_value.
  ENDMETHOD.

ENDCLASS.
