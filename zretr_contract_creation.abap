*&---------------------------------------------------------------------*
*& Report ZRETR_CONTRACT_CREATION
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zretr_contract_creation.


INCLUDE zretr_contract_creation_top.
INCLUDE zretr_contract_creation_cls.
INCLUDE zretr_contract_creation_f01.


INITIALIZATION.
  CREATE OBJECT go_class.

AT SELECTION-SCREEN.
  go_class->ss_ucomm( ).

AT SELECTION-SCREEN OUTPUT.
  go_class->ss_button( ).

*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  go_class->get_local( IMPORTING ev_file_loc = p_file ).


START-OF-SELECTION.
  go_class->call_iref( ).
  IF go_class->get_excel( ) EQ 0.
    CALL SCREEN 9000.
  ENDIF.
