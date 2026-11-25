*&---------------------------------------------------------------------*
*& Include          ZMME004_TOP
*&---------------------------------------------------------------------*
TABLES sscrfields.

TYPES: BEGIN OF ty_char_value,
         charact TYPE atnam,
         value   TYPE string,
       END OF ty_char_value,
       tt_char_value TYPE STANDARD TABLE OF ty_char_value WITH EMPTY KEY.

TYPES: BEGIN OF ty_batch_row,
         icon   TYPE icon_d,
         matnr  TYPE matnr,
         werks  TYPE werks_d,
         charg  TYPE charg_d,
         values TYPE tt_char_value,
         status TYPE string,
       END OF ty_batch_row,
       tt_batch_row TYPE STANDARD TABLE OF ty_batch_row WITH EMPTY KEY.

TYPES: BEGIN OF ty_result,
         ok     TYPE abap_bool,
         status TYPE string,
       END OF ty_result.

TYPES tt_atnam TYPE STANDARD TABLE OF atnam WITH EMPTY KEY.

DATA gv_charg TYPE charg_d.
