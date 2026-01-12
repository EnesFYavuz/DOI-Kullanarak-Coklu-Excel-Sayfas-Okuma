*&---------------------------------------------------------------------*
*& Include          ZRETR_CONTRACT_CREATION_CLS
*&---------------------------------------------------------------------*
CLASS lcl_class DEFINITION.

  PUBLIC SECTION.

    METHODS:
      display_alv,
      get_excel  RETURNING VALUE(rv_subrc) TYPE sy-subrc,
      get_local EXPORTING ev_file_loc TYPE localfile,
      get_sheet_data IMPORTING iv_sheet_name TYPE soi_field_name  it_data TYPE soi_generic_table,
      ss_button,
      ss_ucomm,
      call_iref,
      download_excel_tmp IMPORTING  iv_mime_path TYPE string,
      popup_alv IMPORTING ""iv_st_name TYPE tabname
                  it_table TYPE STANDARD TABLE,
      save IMPORTING iv_test TYPE xfeld,
      activate IMPORTING iv_test TYPE xfeld.

  PRIVATE SECTION.
    DATA: gr_container TYPE REF TO cl_gui_custom_container,
          gr_grid      TYPE REF TO cl_gui_alv_grid,

          gt_fieldcat  TYPE lvc_t_fcat,
          gs_fieldcat  TYPE lvc_t_fcat,
          gs_layout    TYPE lvc_s_layo.
    DATA: lt_return TYPE bapiret2_t.
    METHODS:
      handle_toolbar      FOR EVENT toolbar       OF cl_gui_alv_grid IMPORTING e_object e_interactive,
      user_command        FOR EVENT user_command  OF cl_gui_alv_grid IMPORTING e_ucomm sender,
      on_hotspot_click    FOR EVENT hotspot_click OF cl_gui_alv_grid IMPORTING e_row_id e_column_id es_row_no.
*      handle_on_f4        FOR EVENT onf4          OF cl_gui_alv_grid IMPORTING e_fieldname es_row_no er_event_data,
*      handle_data_changed FOR EVENT data_changed  OF cl_gui_alv_grid IMPORTING er_data_changed
*                                                                               e_onf4 e_onf4_before e_onf4_after.

ENDCLASS.

CLASS lcl_class IMPLEMENTATION.
  METHOD display_alv.

    IF gr_grid IS NOT BOUND.

      CREATE OBJECT gr_container
        EXPORTING
          container_name              = 'GV_CONTAINER'
        EXCEPTIONS
          cntl_error                  = 1
          cntl_system_error           = 2
          create_error                = 3
          lifetime_error              = 4
          lifetime_dynpro_dynpro_link = 5
          OTHERS                      = 6.

      IF gr_container IS BOUND.
        CREATE OBJECT gr_grid
          EXPORTING
            i_parent          = gr_container
          EXCEPTIONS
            error_cntl_create = 1
            error_cntl_init   = 2
            error_cntl_link   = 3
            error_dp_create   = 4
            OTHERS            = 5.
      ENDIF.

      CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
        EXPORTING
          i_structure_name       = 'ZRETR_S_CONTRACT_CREATION'
        CHANGING
          ct_fieldcat            = gt_fieldcat
        EXCEPTIONS
          inconsistent_interface = 1
          program_error          = 2
          OTHERS                 = 3.

      gs_layout-zebra      = abap_true.
      gs_layout-col_opt    = abap_true.
      gs_layout-sel_mode   = 'A'.

      LOOP AT gt_fieldcat ASSIGNING FIELD-SYMBOL(<fs_fieldcat>).

        CASE <fs_fieldcat>-fieldname.
          WHEN 'ICON_S'.
            <fs_fieldcat>-hotspot = abap_true.
            <fs_fieldcat>-col_pos = 0.
            <fs_fieldcat>-scrtext_s = 'Kaydetme Durumu'.
            <fs_fieldcat>-scrtext_m = 'Kaydetme Durumu'.
            <fs_fieldcat>-scrtext_l = 'Kaydetme Durumu'.
            <fs_fieldcat>-reptext   = 'Kaydetme Durumu'.
            <fs_fieldcat>-seltext   = 'Kaydetme Durumu'.
          WHEN 'ICON_A'.
            <fs_fieldcat>-hotspot = abap_true.
            <fs_fieldcat>-col_pos = 1.
            <fs_fieldcat>-scrtext_s = 'Etkinleþtirme  Durumu'.
            <fs_fieldcat>-scrtext_m = 'Etkinleþtirme  Durumu'.
            <fs_fieldcat>-scrtext_l = 'Etkinleþtirme  Durumu'.
            <fs_fieldcat>-reptext   = 'Etkinleþtirme  Durumu'.
            <fs_fieldcat>-seltext   = 'Etkinleþtirme  Durumu'.
          WHEN 'UNIQ_ID'.
            <fs_fieldcat>-hotspot = abap_true.
            <fs_fieldcat>-scrtext_s = 'Uniq ID'.
            <fs_fieldcat>-scrtext_m = 'Sözleþme Uniq ID'.
            <fs_fieldcat>-scrtext_l = 'Sözleþme Uniq ID'.
            <fs_fieldcat>-reptext   = 'Sözleþme Uniq ID'.
            <fs_fieldcat>-seltext   = 'Sözleþme Uniq ID'.

          WHEN 'PARTNER'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Muhatap'.
            <fs_fieldcat>-scrtext_m = 'Muhatap'.
            <fs_fieldcat>-scrtext_l = 'Muhatap'.
            <fs_fieldcat>-reptext   = 'Muhatap'.
            <fs_fieldcat>-seltext   = 'Muhatap'.

          WHEN 'OBJECT_REL'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Sözleþme Nesneleri'.
            <fs_fieldcat>-scrtext_m = 'Sözleþme Nesneleri'.
            <fs_fieldcat>-scrtext_l = 'Sözleþme Nesneleri'.
            <fs_fieldcat>-reptext   = 'Sözleþme Nesneleri'.
            <fs_fieldcat>-seltext   = 'Sözleþme Nesneleri'.

          WHEN 'TERM_PAYM'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Ödeme Koþullarý'.
            <fs_fieldcat>-scrtext_m = 'Ödeme Koþullarý'.
            <fs_fieldcat>-scrtext_l = 'Ödeme Koþullarý'.
            <fs_fieldcat>-reptext   = 'Ödeme Koþullarý'.
            <fs_fieldcat>-seltext   = 'Ödeme Koþullarý'.

          WHEN 'TERM_RHTYM'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Koþul Sýklýðý'.
            <fs_fieldcat>-scrtext_m = 'Koþul Sýklýðý'.
            <fs_fieldcat>-scrtext_l = 'Koþul Sýklýðý'.
            <fs_fieldcat>-reptext   = 'Koþul Sýklýðý'.
            <fs_fieldcat>-seltext   = 'Koþul Sýklýðý'.

          WHEN 'TERM_ORG_ASS'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Organizasyonel Atama'.
            <fs_fieldcat>-scrtext_m = 'Organizasyonel Atama'.
            <fs_fieldcat>-scrtext_l = 'Organizasyonel Atama'.
            <fs_fieldcat>-reptext   = 'Organizasyonel Atama'.
            <fs_fieldcat>-seltext   = 'Organizasyonel Atama'.

          WHEN 'CONDITION'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Koþul'.
            <fs_fieldcat>-scrtext_m = 'Koþul'.
            <fs_fieldcat>-scrtext_l = 'Koþul'.
            <fs_fieldcat>-reptext   = 'Koþul'.
            <fs_fieldcat>-seltext   = 'Koþul'.
          WHEN 'TERM_EVA'.
            <fs_fieldcat>-hotspot     = abap_true.
            <fs_fieldcat>-scrtext_s   = 'Deðerleme Parametresi'.
            <fs_fieldcat>-scrtext_m   = 'Deðerleme Parametresi'.
            <fs_fieldcat>-scrtext_l   = 'Deðerleme Parametresi'.
            <fs_fieldcat>-reptext     = 'Deðerleme Parametresi'.
            <fs_fieldcat>-seltext     = 'Deðerleme Parametresi'.
          WHEN 'TERM_EVA_CONDITION'.
            <fs_fieldcat>-hotspot   = abap_true.
            <fs_fieldcat>-scrtext_s = 'Deðerleme Koþullarý'.
            <fs_fieldcat>-scrtext_m = 'Deðerleme Koþullarý'.
            <fs_fieldcat>-scrtext_l = 'Deðerleme Koþullarý'.
            <fs_fieldcat>-reptext   = 'Deðerleme Koþullarý'.
            <fs_fieldcat>-seltext   = 'Deðerleme Koþullarý'.
          WHEN 'MESSAGE'.
            <fs_fieldcat>-no_out = abap_true.
          WHEN 'ACTIVE'.
            <fs_fieldcat>-no_out = abap_true.

          WHEN OTHERS.
        ENDCASE.
      ENDLOOP.

      IF gr_grid IS BOUND.

        SET HANDLER user_command        FOR gr_grid.
        SET HANDLER handle_toolbar      FOR gr_grid.
        SET HANDLER on_hotspot_click    FOR gr_grid.
        gr_grid->set_table_for_first_display(
          EXPORTING
            is_layout                     = gs_layout
          CHANGING
            it_outtab                     = gt_data
            it_fieldcatalog               = gt_fieldcat
          EXCEPTIONS
            invalid_parameter_combination = 1
            program_error                 = 2
            too_many_lines                = 3
            OTHERS                        = 4 ).

      ENDIF.

    ELSE.
      gr_grid->refresh_table_display( ).
    ENDIF.
  ENDMETHOD.
  METHOD user_command.
    CASE e_ucomm.
      WHEN 'SAVE_SIMU'.
        save( iv_test = 'X' ).
      WHEN 'SAVE'.
        save( iv_test = '' ).
      WHEN 'ACT_SIMU'.
        activate( iv_test = 'X' ).
      WHEN 'ACT'.
        activate( iv_test = '' ).
      WHEN OTHERS.
    ENDCASE.
    IF gt_message IS NOT INITIAL.
      CALL FUNCTION 'FINB_BAPIRET2_DISPLAY'
        EXPORTING
          it_message = gt_message.
      CLEAR gt_message.
    ENDIF.
    gr_grid->refresh_table_display( ).

  ENDMETHOD.
  METHOD get_local.

    CALL FUNCTION 'F4_FILENAME'
      EXPORTING
        program_name  = sy-cprog
        dynpro_number = sy-dynnr
      IMPORTING
        file_name     = ev_file_loc.

  ENDMETHOD.

  METHOD handle_toolbar.
    DATA: ls_button TYPE stb_button.

    CLEAR ls_button.
    ls_button-function  = 'SAVE_SIMU'.
    ls_button-text      = 'Kaydetme Simülasyonu'.
    ls_button-icon      = icon_simulate.
    ls_button-quickinfo = 'Kaydetme Simülasyonu'.
    ls_button-disabled  = abap_false.
    APPEND ls_button TO e_object->mt_toolbar.
    CLEAR ls_button.
    ls_button-function  = 'SAVE'.
    ls_button-text      = 'Kaydet'.
    ls_button-icon      = icon_system_save.
    ls_button-quickinfo = 'Kaydet'.
    ls_button-disabled  = abap_false.
    APPEND ls_button TO e_object->mt_toolbar.
    CLEAR ls_button.
    ls_button-function  = 'ACT_SIMU'.
    ls_button-text      = 'Etkinleþtirme Simülasyonu'.
    ls_button-icon      = icon_simulate .
    ls_button-quickinfo = 'Etkinleþtirme Simülasyonu'.
    ls_button-disabled  = abap_false.
    APPEND ls_button TO e_object->mt_toolbar.
    CLEAR ls_button.
    ls_button-function  = 'ACT'.
    ls_button-text      = 'Etkinleþtir'.
    ls_button-icon      = icon_activate.
    ls_button-quickinfo = 'Etkinleþtir'.
    ls_button-disabled  = abap_false.
    APPEND ls_button TO e_object->mt_toolbar.

  ENDMETHOD.

  METHOD get_excel.
    DATA: lt_sheets TYPE soi_sheets_table,
          lt_data   TYPE soi_generic_table,
          lt_ranges TYPE soi_range_list,
          lt_range  TYPE soi_dimension_table,
          lv_row    TYPE i VALUE 2,
          lv_rows   TYPE i VALUE 2000,
          lv_col    TYPE i VALUE 1,
          lv_cols   TYPE i VALUE 25.
    DATA(lv_url) = CONV char255( |FILE://{ p_file }| ).
    CALL METHOD iref_document->open_document
      EXPORTING
        document_title = 'Excel'
        document_url   = lv_url
        open_inplace   = 'X'
      IMPORTING
        error          = iref_error.
    IF iref_error IS BOUND AND iref_error->has_failed = 'X'.
      CALL METHOD iref_error->raise_message EXPORTING type = 'E'.
      RETURN.
    ENDIF.

    CALL METHOD iref_document->get_spreadsheet_interface
      EXPORTING
        no_flush        = ' '
      IMPORTING
        sheet_interface = iref_spreadsheet
        error           = iref_error.
    IF iref_error IS BOUND AND iref_error->has_failed = 'X'.
      CALL METHOD iref_error->raise_message EXPORTING type = 'E'.
      RETURN.
    ENDIF.

    CALL METHOD iref_spreadsheet->get_sheets
      EXPORTING
        no_flush = ' '
      IMPORTING
        sheets   = lt_sheets
        error    = iref_error.
    IF iref_error IS BOUND AND iref_error->has_failed = 'X'.
      CALL METHOD iref_error->raise_message EXPORTING type = 'E'.
      RETURN.
    ENDIF.

    APPEND VALUE soi_dimension_item(    row     =  lv_row
                                        column  =  lv_col
                                        rows    =  lv_rows
                                        columns =  lv_cols  ) TO lt_range.



    LOOP AT lt_sheets INTO DATA(lv_sheet).

      CALL METHOD iref_spreadsheet->select_sheet
        EXPORTING
          name  = lv_sheet-sheet_name
        IMPORTING
          error = iref_error.

*      CALL METHOD iref_spreadsheet->set_selection
*        EXPORTING
*          top     = lv_row
*          left    = lv_col
*          rows    = lv_rows
*          columns = lv_cols.
*
*      CALL METHOD iref_spreadsheet->insert_range
*        EXPORTING
*          name     = 'Test'
*          rows     = lv_rows
*          columns  = lv_cols
*          no_flush = ''
*        IMPORTING
*          error    = iref_error.
*      IF iref_error->has_failed = 'X'.
*        EXIT.
*      ENDIF.

      REFRESH lt_data.
      CALL METHOD iref_spreadsheet->get_ranges_data
        EXPORTING
          rangesdef = lt_range
        IMPORTING
          contents  = lt_data
          error     = iref_error
        CHANGING
          ranges    = lt_ranges.

      " Boþ satýrlarý at
      DELETE lt_data WHERE value IS INITIAL OR value = space.

      get_sheet_data( iv_sheet_name = lv_sheet-sheet_name it_data = lt_data ).
    ENDLOOP.

  ENDMETHOD.
  METHOD ss_button.
    sscrfields-functxt_01 = TEXT-s01.
    sscrfields-functxt_02 = TEXT-s02.
    sscrfields-functxt_03 = TEXT-s03.
  ENDMETHOD.
  METHOD ss_ucomm.
    CASE sy-ucomm.
      WHEN 'ONLI'.
        IF p_file IS INITIAL.
          MESSAGE 'File must be obligatory' TYPE 'E'.
        ENDIF.
      WHEN 'FC01'.
        download_excel_tmp( gc_fc01 ) .
      WHEN 'FC02'.
        download_excel_tmp( gc_fc02 ) .
      WHEN 'FC03'.
        download_excel_tmp( gc_fc03 ) .
      WHEN OTHERS.
    ENDCASE.
  ENDMETHOD.
  METHOD call_iref.
    CALL METHOD c_oi_container_control_creator=>get_container_control
      IMPORTING
        control = iref_control
        error   = iref_error.
    IF iref_error->has_failed = 'X'.
      CALL METHOD iref_error->raise_message
        EXPORTING
          type = 'E'.
    ENDIF.

    CREATE OBJECT iref_container
      EXPORTING
        container_name = 'CONT'.

    IF sy-subrc <> 0.
      MESSAGE e001(00) WITH 'Container oluþturulamadý'.
    ENDIF.

    CALL METHOD iref_control->init_control
      EXPORTING
        inplace_enabled     = 'X'
        r3_application_name = 'EXCEL CONTAINER'
        parent              = iref_container
      IMPORTING
        error               = iref_error.
    IF iref_error IS BOUND AND iref_error->has_failed = 'X'.
      CALL METHOD iref_error->raise_message EXPORTING type = 'E'.
    ENDIF.
    CALL METHOD iref_control->get_document_proxy
      EXPORTING
        document_type  = soi_doctype_excel_sheet
      IMPORTING
        document_proxy = iref_document
        error          = iref_error.
    IF iref_error IS BOUND AND iref_error->has_failed = 'X'.
      CALL METHOD iref_error->raise_message EXPORTING type = 'E'.
    ENDIF.

  ENDMETHOD.
  METHOD get_sheet_data.
    DATA: lt_comp    TYPE abap_compdescr_tab,
          lv_tabname TYPE string,
          lv_row     TYPE char4,
          ls_data    TYPE soi_generic_item.

    FIELD-SYMBOLS:
      <lt_data>   LIKE gt_data,
      <ls_data>   LIKE LINE OF gt_data,
      <lt_s_data> TYPE table,
      <ls_s_data> TYPE any,
      <fs_value>  TYPE any,
      <lv_icon>   TYPE icon_d.

    ASSIGN gt_data TO <lt_data>.
    IF <lt_data> IS NOT ASSIGNED.
      RETURN.
    ENDIF.

    LOOP AT it_data INTO ls_data GROUP BY ls_data-row INTO lv_row.

      READ TABLE it_data INTO DATA(ls_firstcell)
      WITH KEY row = lv_row column = '1'.
      IF sy-subrc EQ 0.
        DATA(lv_uniq_id) = ls_firstcell-value.
      ENDIF.

      READ TABLE <lt_data> ASSIGNING <ls_data>
       WITH KEY uniq_id = lv_uniq_id.
      IF sy-subrc <> 0.
        APPEND INITIAL LINE TO <lt_data> ASSIGNING <ls_data>.
        <ls_data>-uniq_id = lv_uniq_id.
      ENDIF.



      CASE iv_sheet_name.

        WHEN 'IMPORT_PARAMETERS'.
          lt_comp = CAST cl_abap_structdescr(
                      cl_abap_typedescr=>describe_by_data( <ls_data> )
                    )->components.

          DELETE lt_comp WHERE name = 'ICON_S' OR name = 'ICON_A'.

          LOOP AT GROUP lv_row ASSIGNING FIELD-SYMBOL(<ls_cell_p>).
            READ TABLE lt_comp INDEX <ls_cell_p>-column INTO DATA(ls_comp_p).
            IF sy-subrc = 0.
              ASSIGN COMPONENT ls_comp_p-name OF STRUCTURE <ls_data> TO <fs_value>.
              IF <fs_value> IS ASSIGNED.
                DATA(lv_date) = VALUE sy-datum( ).
                IF ls_comp_p-type_kind = cl_abap_typedescr=>typekind_date.
                  CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
                    EXPORTING
                      date_external            = <ls_cell_p>-value
                      accept_initial_date      = 'X'
                    IMPORTING
                      date_internal            = lv_date
                    EXCEPTIONS
                      date_external_is_invalid = 1                " the external date is invalid (not plausible)
                      OTHERS                   = 2.
                  IF sy-subrc <> 0.
                    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
                  ENDIF.
                  <fs_value> = lv_date.
                ELSE.
                  <fs_value> = <ls_cell_p>-value.
                ENDIF.
              ENDIF.
            ENDIF.
          ENDLOOP.

          ASSIGN COMPONENT 'ICON' OF STRUCTURE <ls_data> TO FIELD-SYMBOL(<fs_icon>).
          IF <fs_icon> IS ASSIGNED.
            <fs_icon> = gv_yellow.
          ENDIF.

        WHEN 'PARTNERS'.
          ASSIGN <ls_data>-partner_t TO <lt_s_data>.
          ASSIGN <ls_data>-partner TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.

        WHEN 'OBJECT_REL'.
          ASSIGN <ls_data>-object_rel_t TO <lt_s_data>.
          ASSIGN <ls_data>-object_rel TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.

        WHEN 'TERM_PAYMENT'.
          ASSIGN <ls_data>-term_paym_t TO <lt_s_data>.
          ASSIGN <ls_data>-term_paym TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.

        WHEN 'TERM_RHYTHM'.
          ASSIGN <ls_data>-term_rhtym_t TO <lt_s_data>.
          ASSIGN <ls_data>-term_rhtym TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.

        WHEN 'TERM_ORG_ASSIGNMENT'.
          ASSIGN <ls_data>-term_org_ass_t TO <lt_s_data>.
          ASSIGN <ls_data>-term_org_ass TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.

        WHEN 'CONDITION'.
          ASSIGN <ls_data>-condition_t TO <lt_s_data>.
          ASSIGN <ls_data>-condition TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.
        WHEN 'TERM_EVALUATION'.
          ASSIGN <ls_data>-term_eva_t TO <lt_s_data>.
          ASSIGN <ls_data>-term_eva TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.
        WHEN 'TERM_EVALUATION_CONDITION'.
          ASSIGN <ls_data>-term_eva_condition_t TO <lt_s_data>.
          ASSIGN <ls_data>-term_eva_condition TO <lv_icon>.
          IF <lv_icon> IS ASSIGNED.
            <lv_icon> = icon_table_settings.
          ENDIF.

        WHEN OTHERS.

      ENDCASE.


      IF <lt_s_data> IS ASSIGNED.
        DATA(lo_descr) = CAST cl_abap_tabledescr(
                           cl_abap_typedescr=>describe_by_data( <lt_s_data> )
                         ).
        lt_comp = CAST cl_abap_structdescr( lo_descr->get_table_line_type( ) )->components.


        APPEND INITIAL LINE TO <lt_s_data> ASSIGNING <ls_s_data>.

        LOOP AT GROUP lv_row ASSIGNING FIELD-SYMBOL(<ls_cell>).
          READ TABLE lt_comp INDEX <ls_cell>-column INTO DATA(ls_comp_np).
          IF sy-subrc = 0.
            ASSIGN COMPONENT ls_comp_np-name OF STRUCTURE <ls_s_data> TO <fs_value>.
            IF <fs_value> IS ASSIGNED.
              DATA(lv_date_np) = VALUE sy-datum( ).
              IF ls_comp_np-type_kind = cl_abap_typedescr=>typekind_date.
                CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
                  EXPORTING
                    date_external            = <ls_cell>-value
                    accept_initial_date      = 'X'
                  IMPORTING
                    date_internal            = lv_date_np
                  EXCEPTIONS
                    date_external_is_invalid = 1
                    OTHERS                   = 2.
                IF sy-subrc <> 0.
                  MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                    WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
                ENDIF.
                <fs_value> = lv_date_np.
              ELSE.
                <fs_value> = <ls_cell>-value.
              ENDIF.

*              <fs_value> = <ls_cell>-value.
            ENDIF.
          ENDIF.
        ENDLOOP.

      ENDIF.

    ENDLOOP.
  ENDMETHOD.
  METHOD download_excel_tmp.
*    DATA: gv_mime_path TYPE string VALUE '.xlsx',
    DATA: lv_mime_file TYPE xstring,
          lv_file_path TYPE string,
          lt_bin_tab   TYPE solix_tab.

    DATA: lv_filename   TYPE string,
          lv_path       TYPE string,
          lv_fullpath   TYPE string,
          lv_result     TYPE i,
          lv_path_local TYPE localfile.

    TRY.
        cl_mime_repository_api=>get_api( )->get(
          EXPORTING
            i_url            = iv_mime_path
            i_check_authority = ''
          IMPORTING
            e_content        = lv_mime_file
        ).
      CATCH cx_root INTO DATA(lx_err).
        MESSAGE lx_err->get_text( ) TYPE 'E'.
    ENDTRY.

    IF lv_mime_file IS INITIAL.
      MESSAGE 'MIME Repository dosyasý okunamadý.' TYPE 'E'.
      EXIT.
    ENDIF.

    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer     = lv_mime_file
      TABLES
        binary_tab = lt_bin_tab
      EXCEPTIONS
        failed     = 1
        OTHERS     = 2.

    IF sy-subrc <> 0.
      MESSAGE 'XSTRING dönüþümünde hata.' TYPE 'E'.
    ENDIF.

    DATA(lv_last_part) = substring_after( val = iv_mime_path sub = '/' occ = -1 ).

    DATA(lv_d_filename)  = substring_before( val = lv_last_part sub = '.' ).

    CALL METHOD cl_gui_frontend_services=>file_save_dialog
      EXPORTING
        default_extension = 'XLS'
        default_file_name = lv_d_filename
        initial_directory = 'C:\temp\'
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = lv_fullpath
        user_action       = lv_result.
    IF lv_result EQ cl_gui_frontend_services=>action_ok.

      CALL FUNCTION 'GUI_DOWNLOAD'
        EXPORTING
          bin_filesize = xstrlen( lv_mime_file )
          filename     = lv_fullpath
          filetype     = 'BIN'
        TABLES
          data_tab     = lt_bin_tab
        EXCEPTIONS
          OTHERS       = 1.

      IF sy-subrc = 0.
*      MESSAGE |Dosya baþarýyla indirildi: { lv_file_path }| TYPE 'S'.
        MESSAGE s004(zre_tr) WITH lv_file_path.
      ELSE.
*      MESSAGE 'Dosya indirme sýrasýnda hata oluþtu.' TYPE 'E'.
        MESSAGE e004(zre_tr).
      ENDIF.
    ELSE.
      MESSAGE e006(zre_tr).
    ENDIF.


  ENDMETHOD.
  METHOD on_hotspot_click.
    FIELD-SYMBOLS <lt_table> TYPE STANDARD TABLE.

    READ TABLE gt_data INTO DATA(ls_data) INDEX es_row_no-row_id.
    IF  sy-subrc = 0.
      CASE e_column_id-fieldname.
        WHEN 'PARTNER'.
          ASSIGN ls_data-partner_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL.
            popup_alv(  it_table = <lt_table> ).
          ENDIF.
        WHEN 'OBJECT_REL'.
          ASSIGN ls_data-object_rel_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv( it_table = <lt_table> ).
          ENDIF.

        WHEN 'TERM_PAYM'.
          ASSIGN ls_data-term_paym_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv(  it_table = <lt_table> ).
          ENDIF.

        WHEN 'TERM_RHTYM'.
          ASSIGN ls_data-term_rhtym_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv( it_table = <lt_table> ).
          ENDIF.

        WHEN 'TERM_ORG_ASS'.
          ASSIGN ls_data-term_org_ass_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv( it_table = <lt_table> ).
          ENDIF.
        WHEN 'CONDITION'.
          ASSIGN ls_data-condition_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv(  it_table = <lt_table> ).
          ENDIF.
        WHEN 'TERM_EVA'.
          ASSIGN ls_data-term_eva_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv(  it_table = <lt_table> ).
          ENDIF.
        WHEN 'TERM_EVA_CONDITION'.
          ASSIGN ls_data-term_eva_condition_t TO <lt_table>.
          IF <lt_table> IS ASSIGNED AND <lt_table> IS NOT INITIAL..
            popup_alv(  it_table = <lt_table> ).
          ENDIF.
        WHEN 'ICON_S'.
          CALL FUNCTION 'FINB_BAPIRET2_DISPLAY'
            EXPORTING
              it_message = ls_data-message_s.
        WHEN 'ICON_A'.
          CALL FUNCTION 'FINB_BAPIRET2_DISPLAY'
            EXPORTING
              it_message = ls_data-message_a.

      ENDCASE.
    ENDIF.

  ENDMETHOD.
  METHOD popup_alv.
    DATA: lt_fieldcat TYPE slis_t_fieldcat_alv,
          ls_fieldcat TYPE slis_fieldcat_alv,
          ls_private  TYPE slis_data_caller_exit.

    DATA: lo_structdescr TYPE REF TO cl_abap_structdescr,
          lt_components  TYPE cl_abap_structdescr=>component_table.


    lo_structdescr ?= cl_abap_typedescr=>describe_by_data_ref( REF #( it_table[ 1 ] ) ).
    lt_components = lo_structdescr->get_components( ).

    LOOP AT lt_components INTO DATA(ls_comp).

      CLEAR ls_fieldcat.

      ls_fieldcat-fieldname = ls_comp-name.
      ls_fieldcat-seltext_m = ls_comp-name.
      ls_fieldcat-seltext_l = ls_comp-name.
      ls_fieldcat-outputlen = ls_comp-type->length.
      ls_fieldcat-datatype  = ls_comp-type->type_kind.

      APPEND ls_fieldcat TO lt_fieldcat.

    ENDLOOP.

    ls_private-columnopt = 'X'.

    CALL FUNCTION 'REUSE_ALV_POPUP_TO_SELECT'
      EXPORTING
*       i_title               = ''
        i_allow_no_selection  = 'X'
        i_tabname             = 'GS_RETURN_TAB'
*       i_structure_name      = iv_st_name
        it_fieldcat           = lt_fieldcat
        is_private            = ls_private
        i_screen_start_column = 10      " Soldan uzaklýk
        i_screen_start_line   = 5       " Yukarýdan uzaklýk
        i_screen_end_column   = 200     " Sað sýnýr
        i_screen_end_line     = 25      " Alt sýnýr
      TABLES
        t_outtab              = it_table.
  ENDMETHOD.
  METHOD save.
    DATA: lt_rows TYPE lvc_t_row.
    DATA: lt_term_org_assignment TYPE TABLE OF bapi_re_term_oa_dat,
          lt_term_payment        TYPE TABLE OF bapi_re_term_py_dat,
          lt_term_rhythm         TYPE TABLE OF bapi_re_term_rh_dat,
          lt_partner             TYPE TABLE OF bapi_re_partner_dat,
          lt_object_rel          TYPE TABLE OF bapi_re_object_rel_dat,
          lt_condition           TYPE TABLE OF bapi_re_condition_dat,
          lt_term_eva            TYPE TABLE OF bapi_re_term_ce_dat,
          lt_term_eva_cond       TYPE TABLE OF bapi_re_term_cecond_dat.
    CALL METHOD gr_grid->get_selected_rows
      IMPORTING
        et_index_rows = lt_rows.
    IF lt_rows[] IS  INITIAL.
      MESSAGE s003(zre_tr) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    LOOP AT lt_rows INTO DATA(ls_row).
      READ TABLE gt_data ASSIGNING FIELD-SYMBOL(<fs_data>) INDEX ls_row-index .
      IF sy-subrc EQ 0.
*        IF  <fs_data>-contract_number IS INITIAL.

        IF <fs_data>-contract_number IS NOT INITIAL.
          APPEND VALUE #( type       = 'I'
                          id         = 'ZRE_TR'
                          number     = '002'
                          message_v1 = <fs_data>-uniq_id ) TO gt_message.
        ELSE.

          lt_term_org_assignment = CORRESPONDING #( <fs_data>-term_org_ass_t ).
          lt_term_payment        = CORRESPONDING #( <fs_data>-term_paym_t ).
          lt_term_rhythm         = CORRESPONDING #( <fs_data>-term_rhtym_t ).
          lt_partner             = CORRESPONDING #( <fs_data>-partner_t ).
          lt_object_rel          = CORRESPONDING #( <fs_data>-object_rel_t ).
          lt_condition           = CORRESPONDING #( <fs_data>-condition_t ).
          lt_term_eva            = CORRESPONDING #( <fs_data>-term_eva_t ).
          lt_term_eva_cond       = CORRESPONDING #( <fs_data>-term_eva_condition_t ).

          CALL FUNCTION 'BAPI_RE_CN_CREATE'
            EXPORTING
              comp_code_ext             = <fs_data>-comp_code_ext
              contract_type             = <fs_data>-contract_type
              contract                  = CORRESPONDING bapi_re_contract_dat( <fs_data> )
              test_run                  = iv_test
            IMPORTING
              contractnumber            = <fs_data>-contract_number
            TABLES
              term_org_assignment       = lt_term_org_assignment
              term_payment              = lt_term_payment
              term_rhythm               = lt_term_rhythm
              partner                   = lt_partner
              object_rel                = lt_object_rel
              condition                 = lt_condition
              term_evaluation           = lt_term_eva
              term_evaluation_condition = lt_term_eva_cond
              return                    = lt_return.
          IF sy-subrc EQ 0.
            <fs_data>-message_s = lt_return.
            DATA(lv_error) =  xsdbool(  line_exists( lt_return[ type = 'E' ] )  OR line_exists( lt_return[ type = 'A' ] )  OR line_exists( lt_return[ type = 'X' ] ) ).
            IF iv_test IS INITIAL.
              IF lv_error = abap_true.
                CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
                <fs_data>-icon_s = gv_red.
              ELSE.
                CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
                  EXPORTING
                    wait = abap_true.
                <fs_data>-icon_s = gv_green.

              ENDIF.
            ELSE.""                           Simülasyon
              IF lv_error = abap_true.
                <fs_data>-icon_s = gv_red.
              ELSE.
                <fs_data>-icon_s = gv_yellow.
              ENDIF.
            ENDIF.
          ENDIF.
          CLEAR: lt_term_org_assignment,lt_term_payment,lt_term_rhythm,lt_partner,
                 lt_object_rel,lt_condition,lt_term_eva,lt_term_eva_cond,lt_return.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.
  METHOD activate.
    DATA: lt_rows TYPE lvc_t_row.
    DATA: lt_term_org_assignment_c TYPE TABLE OF bapi_re_term_oa_datc,
          lt_term_payment_c        TYPE TABLE OF bapi_re_term_py_datc,
          lt_term_rhythm_c         TYPE TABLE OF bapi_re_term_rh_datc,
          lt_partner_c             TYPE TABLE OF bapi_re_partner_datc,
          lt_object_rel_c          TYPE TABLE OF bapi_re_object_rel_datc,
          lt_condition_c           TYPE TABLE OF bapi_re_condition_datc,
          lt_term_eva_c            TYPE TABLE OF bapi_re_term_ce_datc,
          lt_term_eva_cond_c       TYPE TABLE OF bapi_re_term_cecond_datc.
    CALL METHOD gr_grid->get_selected_rows
      IMPORTING
        et_index_rows = lt_rows.
    IF lt_rows[] IS  INITIAL.
      MESSAGE s003(zre_tr) DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.
    LOOP AT lt_rows INTO DATA(ls_row).
      READ TABLE gt_data ASSIGNING FIELD-SYMBOL(<fs_data>) INDEX ls_row-index .
      IF sy-subrc EQ 0.
*        IF <fs_data>-contract_number IS NOT INITIAL AND <fs_data>-icon eq IS INITIAL.
        IF <fs_data>-contract_number IS INITIAL.
          ""kaydedilmeden etkinleþtirilemez
          APPEND VALUE #(   type       = 'E'
                            id         = 'ZRE_TR'
                            number     = '000'
                            message_v1 = <fs_data>-uniq_id        ) TO gt_message.


        ELSEIF <fs_data>-active IS NOT INITIAL.
          ""satýr daha önce aktifleþtirilmiþtir
          APPEND VALUE #( type       = 'I'
                          id         = 'ZRE_TR'
                          number     = '001'
                          message_v1 = <fs_data>-uniq_id         ) TO gt_message.
        ELSE.

          lt_term_org_assignment_c = CORRESPONDING #( <fs_data>-term_org_ass_t ).
          lt_term_payment_c        = CORRESPONDING #( <fs_data>-term_paym_t ).
          lt_term_rhythm_c         = CORRESPONDING #( <fs_data>-term_rhtym_t ).
          lt_partner_c             = CORRESPONDING #( <fs_data>-partner_t ).
          lt_object_rel_c          = CORRESPONDING #( <fs_data>-object_rel_t ).
          lt_condition_c           = CORRESPONDING #( <fs_data>-condition_t ).
          lt_term_eva_c            = CORRESPONDING #( <fs_data>-term_eva_t ).
          lt_term_eva_cond_c       = CORRESPONDING #( <fs_data>-term_eva_condition_t ).

          CALL FUNCTION 'BAPI_RE_CN_CHANGE'
            EXPORTING
              compcode                  = <fs_data>-comp_code_ext
              contractnumber            = <fs_data>-contract_number
              contract                  = CORRESPONDING bapi_re_contract_dat( <fs_data> )
              test_run                  = iv_test
              trans                     = 'MCAK'
            TABLES
              term_org_assignment       = lt_term_org_assignment_c
              term_payment              = lt_term_payment_c
              term_rhythm               = lt_term_rhythm_c
              partner                   = lt_partner_c
              object_rel                = lt_object_rel_c
              condition                 = lt_condition_c
              term_evaluation           = lt_term_eva_c
              term_evaluation_condition = lt_term_eva_cond_c
              return                    = lt_return.
          IF sy-subrc EQ 0.
            <fs_data>-message_a = lt_return.
            DATA(lv_error) =  xsdbool(  line_exists( lt_return[ type = 'E' ] )  OR line_exists( lt_return[ type = 'A' ] )  OR line_exists( lt_return[ type = 'X' ] ) ).
            IF iv_test IS INITIAL.
              IF lv_error = abap_true.
                CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
                <fs_data>-icon_a = gv_red.
              ELSE.
                CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
                  EXPORTING
                    wait = abap_true.
                <fs_data>-active = abap_true.
                <fs_data>-icon_a = gv_green.
              ENDIF.
            ELSE.                                                                                         ""Simülasyon
              IF lv_error = abap_true.
                <fs_data>-icon_a = gv_red.
              ELSE.
                <fs_data>-icon_a = gv_yellow.
              ENDIF.
            ENDIF.
          ENDIF.
          CLEAR: lt_term_org_assignment_c,lt_term_payment_c,lt_term_rhythm_c,lt_partner_c,
                 lt_object_rel_c,lt_condition_c,lt_term_eva_c,lt_term_eva_cond_c,lt_return.
        ENDIF.
      ENDIF.
    ENDLOOP.





  ENDMETHOD.

ENDCLASS.
