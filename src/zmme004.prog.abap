*&---------------------------------------------------------------------*
*& Report ZMME004
*&---------------------------------------------------------------------*
* Project      : SAP S/4 HANA Implementation
* Program Name : ZMME004
* Description  : Update Batch Characteristic
* Create By    : Kachain N.
* Create Date  : 25.11.2026
* Transport No.: S4DK900108
*----------------------------------------------------------------------*
* Modification History
*----------------------------------------------------------------------*
* Change Date  : DD MMM YYYY
* Change By    : XXXXX X.
* Transport No.: XXXXXXXXX
* Ticket Number: PRJXX/PRJTASKXXX
* Search Term  : CHXXX
* Description  :
*----------------------------------------------------------------------*
REPORT zmme004 NO STANDARD PAGE HEADING.

*======================================================================*
*----------------------INCLUDE-----------------------------------------*
*======================================================================*
INCLUDE :
  zmme004_top ,"Global Variable
  zmme004_s01 ,"Selection Screen
  zmme004_d01 ,"Local Class Definition
  zmme004_p01 ."Local Class Implementation

*======================================================================*
*----------------------INITIALIZATION ---------------------------------*
*======================================================================*
INITIALIZATION .
  lcl_screen=>init( ).

*======================================================================*
*----------------------AT SELECTION-SCREEN OUTPUT----------------------*
*======================================================================*
AT SELECTION-SCREEN OUTPUT .
  lcl_screen=>pbo( ).

*======================================================================*
*----------------------AT SELECTION-SCREEN ----------------------------*
*======================================================================*
AT SELECTION-SCREEN .
  lcl_screen=>pai( ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  lcl_screen=>f4_file( ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_atn1.
  lcl_screen=>f4_characteristic( 'P_ATN1' ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_atn2.
  lcl_screen=>f4_characteristic( 'P_ATN2' ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_atn3.
  lcl_screen=>f4_characteristic( 'P_ATN3' ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_atn4.
  lcl_screen=>f4_characteristic( 'P_ATN4' ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_atn5.
  lcl_screen=>f4_characteristic( 'P_ATN5' ).

*======================================================================*
*----------------------START-OF-SELECTION------------------------------*
*======================================================================*
START-OF-SELECTION .
  NEW lcl_app( )->run( ).
