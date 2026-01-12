*&---------------------------------------------------------------------*
*& Include          ZRETR_CONTRACT_CREATION_TOP
*&---------------------------------------------------------------------*

TABLES: sscrfields.
SELECTION-SCREEN FUNCTION KEY 1.
SELECTION-SCREEN FUNCTION KEY 2.
SELECTION-SCREEN FUNCTION KEY 3.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS p_file TYPE localfile.
SELECTION-SCREEN END OF BLOCK b1.

CLASS lcl_class DEFINITION DEFERRED.

DATA: go_class TYPE REF TO lcl_class,
      ok_code  TYPE sy-ucomm,
      gv_code  TYPE sy-ucomm.

DATA: gt_data    TYPE TABLE OF zretr_s_contract_creation,
      gs_data    TYPE zretr_s_contract_creation,
      gt_message TYPE bapiret2_t.

DATA: gv_filename TYPE string.

CONSTANTS: gv_red    TYPE icon_d VALUE '@0A@',
           gv_yellow TYPE icon_d VALUE '@09@',
           gv_green  TYPE icon_d VALUE '@08@'.

DATA: iref_control     TYPE REF TO i_oi_container_control,
      iref_container   TYPE REF TO cl_gui_custom_container,
      iref_document    TYPE REF TO i_oi_document_proxy,
      iref_spreadsheet TYPE REF TO i_oi_spreadsheet,
      iref_error       TYPE REF TO i_oi_error.

CONSTANTS: gc_fc01 TYPE string VALUE '/SAP/PUBLIC/ZRETR_CONTRACT_CREATION/1Excel Upload Sablonu.xlsx',
           gc_fc02 TYPE string VALUE '/SAP/PUBLIC/ZRETR_CONTRACT_CREATION/2Excel Upload Sablonu - Aciklamali.xlsx',
           gc_fc03 TYPE string VALUE '/SAP/PUBLIC/ZRETR_CONTRACT_CREATION/3IFRS16 Ornek Excel Upload Dosyasi.xlsx'.
