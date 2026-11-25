*&---------------------------------------------------------------------*
*& Include          ZMME004_D01
*&---------------------------------------------------------------------*
CLASS lcx_error DEFINITION INHERITING FROM cx_static_check FINAL.
  PUBLIC SECTION.
    METHODS constructor
      IMPORTING iv_text TYPE string.

    METHODS if_message~get_text REDEFINITION.

  PRIVATE SECTION.
    DATA mv_text TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Excel workbook (.xlsx) from the front end
*&---------------------------------------------------------------------*
CLASS lcl_workbook DEFINITION FINAL.
  PUBLIC SECTION.
    TYPES tt_sheet TYPE STANDARD TABLE OF string_table WITH EMPTY KEY.

    CLASS-METHODS is_open
      IMPORTING iv_file        TYPE string
      RETURNING VALUE(rv_open) TYPE abap_bool.

    CLASS-METHODS sheet_names
      IMPORTING iv_file         TYPE string
      RETURNING VALUE(rt_names) TYPE string_table
      RAISING   lcx_error.

    CLASS-METHODS read_sheet
      IMPORTING iv_file        TYPE string
                iv_sheet       TYPE string
      RETURNING VALUE(rt_rows) TYPE tt_sheet
      RAISING   lcx_error.

  PRIVATE SECTION.
    CLASS-DATA: gv_file  TYPE string,
                go_excel TYPE REF TO cl_fdt_xl_spreadsheet.

    CLASS-METHODS open
      IMPORTING iv_file         TYPE string
      RETURNING VALUE(ro_excel) TYPE REF TO cl_fdt_xl_spreadsheet
      RAISING   lcx_error.

    CLASS-METHODS upload
      IMPORTING iv_file           TYPE string
      RETURNING VALUE(rv_xstring) TYPE xstring
      RAISING   lcx_error.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Batch classification (class type 023)
*&---------------------------------------------------------------------*
CLASS lcl_batch_char DEFINITION FINAL.
  PUBLIC SECTION.
    TYPES: BEGIN OF ty_char_text,
             atnam TYPE atnam,
             atbez TYPE atbez,
           END OF ty_char_text,
           tt_char_text TYPE STANDARD TABLE OF ty_char_text WITH EMPTY KEY.

    METHODS batch_class
      IMPORTING iv_matnr        TYPE matnr
      RETURNING VALUE(rv_class) TYPE klasse_d.

    METHODS characteristic_list
      IMPORTING iv_class       TYPE klasse_d OPTIONAL
      RETURNING VALUE(rt_list) TYPE tt_char_text.

    METHODS descriptions
      IMPORTING it_chars       TYPE tt_atnam
      RETURNING VALUE(rt_text) TYPE tt_char_text.

    METHODS update
      IMPORTING iv_matnr         TYPE matnr
                iv_charg         TYPE charg_d
                it_values        TYPE tt_char_value
      RETURNING VALUE(rs_result) TYPE ty_result.

  PRIVATE SECTION.
    CONSTANTS: c_klart TYPE klah-klart VALUE '023',
               c_blank TYPE string     VALUE 'BLANK'.

    TYPES: BEGIN OF ty_char_def,
             atnam TYPE atnam,
             atfor TYPE atfor,
             atein TYPE atein,
           END OF ty_char_def,
           tt_char_def TYPE SORTED TABLE OF ty_char_def WITH NON-UNIQUE KEY atnam.

    TYPES: BEGIN OF ty_class_cache,
             matnr TYPE matnr,
             class TYPE klasse_d,
           END OF ty_class_cache,
           BEGIN OF ty_char_cache,
             class TYPE klasse_d,
             chars TYPE tt_char_def,
           END OF ty_char_cache.

    TYPES: tt_num  TYPE STANDARD TABLE OF bapi1003_alloc_values_num  WITH EMPTY KEY,
           tt_char TYPE STANDARD TABLE OF bapi1003_alloc_values_char WITH EMPTY KEY,
           tt_curr TYPE STANDARD TABLE OF bapi1003_alloc_values_curr WITH EMPTY KEY,
           BEGIN OF ty_values,
             num  TYPE tt_num,
             char TYPE tt_char,
             curr TYPE tt_curr,
           END OF ty_values.

    DATA: mt_class_cache TYPE HASHED TABLE OF ty_class_cache WITH UNIQUE KEY matnr,
          mt_char_cache  TYPE HASHED TABLE OF ty_char_cache  WITH UNIQUE KEY class.

    METHODS class_chars
      IMPORTING iv_class        TYPE klasse_d
      RETURNING VALUE(rt_chars) TYPE tt_char_def.

    METHODS apply_value
      IMPORTING is_char   TYPE ty_char_def
                iv_value  TYPE string
      CHANGING  cs_values TYPE ty_values
      RAISING   lcx_error.

    METHODS set_char_value
      IMPORTING is_char   TYPE ty_char_def
                iv_value  TYPE string
      CHANGING  ct_char   TYPE tt_char.

    METHODS to_number
      IMPORTING is_char         TYPE ty_char_def
                iv_value        TYPE string
      RETURNING VALUE(rv_value) TYPE f
      RAISING   lcx_error.

    METHODS to_date_text
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_text) TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Selection screen events
*&---------------------------------------------------------------------*
CLASS lcl_screen DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS init.
    CLASS-METHODS pbo.
    CLASS-METHODS pai.
    CLASS-METHODS f4_file.

    CLASS-METHODS f4_characteristic
      IMPORTING iv_field TYPE dynfnam.

    CLASS-METHODS char_values
      RETURNING VALUE(rt_values) TYPE tt_char_value.

  PRIVATE SECTION.
    CLASS-METHODS check_input.

    CLASS-METHODS fill_sheet_list
      RETURNING VALUE(rv_sheet) TYPE string.

    CLASS-METHODS read_screen_field
      IMPORTING iv_field        TYPE dynfnam
      RETURNING VALUE(rv_value) TYPE string.

    CLASS-METHODS update_screen_field
      IMPORTING iv_field TYPE dynfnam
                iv_value TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Main flow: collect rows, update, show result
*&---------------------------------------------------------------------*
CLASS lcl_app DEFINITION FINAL.
  PUBLIC SECTION.
    METHODS run.

    CLASS-METHODS to_matnr
      IMPORTING iv_text         TYPE string
      RETURNING VALUE(rv_matnr) TYPE matnr.

  PRIVATE SECTION.
    CONSTANTS: c_first_char_col TYPE i VALUE 4.

    DATA: mt_chars TYPE tt_atnam,
          mt_rows  TYPE tt_batch_row.

    METHODS read_excel
      RAISING lcx_error.

    METHODS read_screen
      RAISING lcx_error.

    METHODS check_plant_authority
      RETURNING VALUE(rv_ok) TYPE abap_bool.

    METHODS update_batches.

    METHODS show_result.

    METHODS build_output_table
      RETURNING VALUE(rr_table) TYPE REF TO data.

    METHODS set_column_texts
      IMPORTING io_alv TYPE REF TO cl_salv_table
      RAISING   cx_salv_not_found.

    METHODS set_column_text
      IMPORTING io_column TYPE REF TO cl_salv_column
                iv_text   TYPE csequence.

    METHODS column_name
      IMPORTING iv_index       TYPE i
      RETURNING VALUE(rv_name) TYPE lvc_fname.

    METHODS alpha_in
      IMPORTING iv_text  TYPE string
      EXPORTING ev_value TYPE clike.
ENDCLASS.
