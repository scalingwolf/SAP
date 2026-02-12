REPORT z_test13.

"===============================================================
" 1) INPUT TYPE (EDIT THIS STRUCTURE ONLY)
"    Must contain SHEETNAME
"===============================================================
TYPES: BEGIN OF ty_row,
         sheetname TYPE string,
         col_a     TYPE string,
         col_b     TYPE string,
         qty       TYPE i,
         amount    TYPE p LENGTH 12 DECIMALS 2,
       END OF ty_row.
TYPES tt_row TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY.

"===============================================================
" 2) XLSX BUILDER (local class first)
"===============================================================
CLASS lcl_xlsx DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS build_from_itab
      IMPORTING it_data TYPE tt_row
      RETURNING VALUE(rv_xlsx) TYPE xstring.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_sheet,
             name TYPE string,
             data TYPE REF TO tt_row,
           END OF ty_sheet.
    TYPES ty_sheets TYPE STANDARD TABLE OF ty_sheet WITH EMPTY KEY.

    CLASS-METHODS group_sheets
      IMPORTING it_data TYPE tt_row
      RETURNING VALUE(rt_sheets) TYPE ty_sheets.

    CLASS-METHODS sheetname_ok
      IMPORTING iv_name TYPE string
      RETURNING VALUE(rv_name) TYPE string.

    CLASS-METHODS esc_xml
      IMPORTING iv_text TYPE string
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS col_letters
      IMPORTING iv_idx TYPE i
      RETURNING VALUE(rv_col) TYPE string.

    CLASS-METHODS comps_wo_sheet
      IMPORTING ir_str TYPE REF TO cl_abap_structdescr
      RETURNING VALUE(rt_comp) TYPE abap_component_tab.

    CLASS-METHODS str_to_xstr
      IMPORTING iv_text TYPE string
      RETURNING VALUE(rv_xstr) TYPE xstring.

    CLASS-METHODS mk_content_types
      IMPORTING iv_sheet_count TYPE i
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS mk_root_rels
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS mk_workbook
      IMPORTING it_sheets TYPE ty_sheets
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS mk_workbook_rels
      IMPORTING iv_sheet_count TYPE i
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS mk_styles
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS mk_worksheet
      IMPORTING it_rows TYPE tt_row
      RETURNING VALUE(rv_xml) TYPE string.
ENDCLASS.

CLASS lcl_xlsx IMPLEMENTATION.

  METHOD build_from_itab.
    DATA(lt_sheets) = group_sheets( it_data ).
    IF lt_sheets IS INITIAL.
      MESSAGE 'No sheets / no data' TYPE 'E'.
    ENDIF.

    DATA(lo_zip) = NEW cl_abap_zip( ).

    lo_zip->add( name = '[Content_Types].xml'
                 content = str_to_xstr( mk_content_types( lines( lt_sheets ) ) ) ).
    lo_zip->add( name = '_rels/.rels'
                 content = str_to_xstr( mk_root_rels( ) ) ).
    lo_zip->add( name = 'xl/workbook.xml'
                 content = str_to_xstr( mk_workbook( lt_sheets ) ) ).
    lo_zip->add( name = 'xl/_rels/workbook.xml.rels'
                 content = str_to_xstr( mk_workbook_rels( lines( lt_sheets ) ) ) ).
    lo_zip->add( name = 'xl/styles.xml'
                 content = str_to_xstr( mk_styles( ) ) ).

    DATA lv_idx TYPE i VALUE 0.
    LOOP AT lt_sheets INTO DATA(ls_sheet).
      lv_idx += 1.
      lo_zip->add(
        name    = |xl/worksheets/sheet{ lv_idx }.xml|
        content = str_to_xstr( mk_worksheet( ls_sheet-data->* ) ) ).
    ENDLOOP.

    rv_xlsx = lo_zip->save( ).
  ENDMETHOD.

  METHOD group_sheets.
    DATA lt TYPE tt_row.
    lt = it_data.

    SORT lt BY sheetname.

    DATA lr_bucket TYPE REF TO tt_row.
    CREATE DATA lr_bucket.
    CLEAR lr_bucket->*.

    DATA lv_prev TYPE string.
    DATA lv_curr TYPE string.

    LOOP AT lt INTO DATA(ls_row).
      lv_curr = sheetname_ok( ls_row-sheetname ).
      IF lv_prev IS INITIAL.
        lv_prev = lv_curr.
      ENDIF.

      IF lv_curr <> lv_prev.
        APPEND VALUE ty_sheet( name = lv_prev data = lr_bucket ) TO rt_sheets.
        CREATE DATA lr_bucket.
        CLEAR lr_bucket->*.
        lv_prev = lv_curr.
      ENDIF.

      APPEND ls_row TO lr_bucket->*.
    ENDLOOP.

    IF lv_prev IS NOT INITIAL.
      APPEND VALUE ty_sheet( name = lv_prev data = lr_bucket ) TO rt_sheets.
    ENDIF.
  ENDMETHOD.

  METHOD mk_content_types.
    DATA(lv) =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">| &&
      |<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>| &&
      |<Default Extension="xml" ContentType="application/xml"/>| &&
      |<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>| &&
      |<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>|.

    DO iv_sheet_count TIMES.
      lv &&= |<Override PartName="/xl/worksheets/sheet{ sy-index }.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>|.
    ENDDO.

    lv &&= |</Types>|.
    rv_xml = lv.
  ENDMETHOD.

  METHOD mk_root_rels.
    rv_xml =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">| &&
      |<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>| &&
      |</Relationships>|.
  ENDMETHOD.

  METHOD mk_workbook.
    DATA(lv) =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">| &&
      |<sheets>|.

    DATA lv_i TYPE i VALUE 0.
    LOOP AT it_sheets INTO DATA(ls).
      lv_i += 1.
      lv &&= |<sheet name="{ esc_xml( ls-name ) }" sheetId="{ lv_i }" r:id="rId{ lv_i }"/>|.
    ENDLOOP.

    lv &&= |</sheets></workbook>|.
    rv_xml = lv.
  ENDMETHOD.

  METHOD mk_workbook_rels.
    DATA(lv) =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">|.

    DO iv_sheet_count TIMES.
      lv &&= |<Relationship Id="rId{ sy-index }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{ sy-index }.xml"/>|.
    ENDDO.

    lv &&= |<Relationship Id="rId{ iv_sheet_count + 1 }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>|.
    lv &&= |</Relationships>|.
    rv_xml = lv.
  ENDMETHOD.

  METHOD mk_styles.
    rv_xml =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">| &&
      |<fonts count="2">| &&
      |<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>| &&
      |<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>| &&
      |</fonts>| &&
      |<fills count="1"><fill><patternFill patternType="none"/></fill></fills>| &&
      |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>| &&
      |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>| &&
      |<cellXfs count="2">| &&
      |<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>| &&
      |<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>| &&
      |</cellXfs>| &&
      |</styleSheet>|.
  ENDMETHOD.

  METHOD mk_worksheet.
    FIELD-SYMBOLS: <val> TYPE any.

    DATA(lo_str) = CAST cl_abap_structdescr( cl_abap_typedescr=>describe_by_data( VALUE ty_row( ) ) ).
    DATA(lt_comp) = comps_wo_sheet( lo_str ).

    DATA(lv) =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">| &&
      |<sheetData>|.

    "Header
    lv &&= |<row r="1">|.
    DATA lv_c TYPE i VALUE 0.
    LOOP AT lt_comp INTO DATA(ls_comp).
      lv_c += 1.
      DATA(lv_ref_h) = |{ col_letters( lv_c ) }1|.
      lv &&= |<c r="{ lv_ref_h }" t="inlineStr" s="1"><is><t>{ esc_xml( ls_comp-name ) }</t></is></c>|.
    ENDLOOP.
    lv &&= |</row>|.

    "Data as text (safe)
    DATA lv_r TYPE i VALUE 1.
    LOOP AT it_rows INTO DATA(ls_row).
      lv_r += 1.
      lv &&= |<row r="{ lv_r }">|.
      lv_c = 0.

      LOOP AT lt_comp INTO ls_comp.
        lv_c += 1.
        DATA(lv_ref) = |{ col_letters( lv_c ) }{ lv_r }|.

        ASSIGN COMPONENT ls_comp-name OF STRUCTURE ls_row TO <val>.
        DATA(lv_text) = COND string( WHEN sy-subrc = 0 THEN |{ <val> }| ELSE `` ).

        lv &&= |<c r="{ lv_ref }" t="inlineStr"><is><t>{ esc_xml( lv_text ) }</t></is></c>|.
      ENDLOOP.

      lv &&= |</row>|.
    ENDLOOP.

    lv &&= |</sheetData></worksheet>|.
    rv_xml = lv.
  ENDMETHOD.

  METHOD comps_wo_sheet.
    rt_comp = ir_str->get_components( ).
    DELETE rt_comp WHERE name = 'SHEETNAME'.
  ENDMETHOD.

  METHOD col_letters.
    DATA n   TYPE i.
    DATA rem TYPE i.
    DATA ch  TYPE c LENGTH 1.

    n = iv_idx.
    rv_col = ``.
    WHILE n > 0.
      rem = ( n - 1 ) MOD 26.
      ch = sy-abcde+rem(1).
      rv_col = |{ ch }{ rv_col }|.
      n = ( n - 1 ) DIV 26.
    ENDWHILE.
  ENDMETHOD.

  METHOD sheetname_ok.
    rv_name = iv_name.
    CONDENSE rv_name.
    REPLACE ALL OCCURRENCES OF ':' IN rv_name WITH '-'.
    REPLACE ALL OCCURRENCES OF '\' IN rv_name WITH '-'.
    REPLACE ALL OCCURRENCES OF '/' IN rv_name WITH '-'.
    REPLACE ALL OCCURRENCES OF '?' IN rv_name WITH ''.
    REPLACE ALL OCCURRENCES OF '*' IN rv_name WITH ''.
    REPLACE ALL OCCURRENCES OF '[' IN rv_name WITH '('.
    REPLACE ALL OCCURRENCES OF ']' IN rv_name WITH ')'.
    IF rv_name IS INITIAL.
      rv_name = 'Sheet'.
    ENDIF.
    IF strlen( rv_name ) > 31.
      rv_name = rv_name+0(31).
    ENDIF.
  ENDMETHOD.

  METHOD esc_xml.
    rv_text = iv_text.
    REPLACE ALL OCCURRENCES OF '&' IN rv_text WITH '&amp;'.
    REPLACE ALL OCCURRENCES OF '<' IN rv_text WITH '&lt;'.
    REPLACE ALL OCCURRENCES OF '>' IN rv_text WITH '&gt;'.
    REPLACE ALL OCCURRENCES OF '"' IN rv_text WITH '&quot;'.
    DATA(lv_apo) = ''''.
    REPLACE ALL OCCURRENCES OF lv_apo IN rv_text WITH '&apos;'.
  ENDMETHOD.

  METHOD str_to_xstr.
    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = iv_text
      IMPORTING
        buffer = rv_xstr
      EXCEPTIONS
        failed = 1
        OTHERS = 2.
    IF sy-subrc <> 0.
      rv_xstr = ``.
    ENDIF.
  ENDMETHOD.

ENDCLASS.

"===============================================================
" 3) Selection + Main
"===============================================================
PARAMETERS: p_subj TYPE so_obj_des OBLIGATORY DEFAULT 'Z_TEST13 XLSX',
            p_file TYPE string     OBLIGATORY DEFAULT 'z_test13.xlsx',
            p_send TYPE abap_bool  AS CHECKBOX DEFAULT abap_true.
SELECT-OPTIONS: so_mail FOR sy-uname NO INTERVALS.

DATA gt_data TYPE tt_row.

START-OF-SELECTION.

  "Demo (replace with your real data)
  gt_data = VALUE tt_row(
    ( sheetname = 'Sales' col_a = 'A1' col_b = 'B1' qty = 10 amount = '123.45' )
    ( sheetname = 'Sales' col_a = 'A2' col_b = 'B2' qty = 20 amount = '50.00'  )
    ( sheetname = 'Stock' col_a = 'S1' col_b = 'W1' qty =  5 amount = '77.70'  )
  ).

  DATA(lv_xlsx) = lcl_xlsx=>build_from_itab( gt_data ).
  WRITE: / |XLSX bytes: { xstrlen( lv_xlsx ) }|.

  IF p_send = abap_true.
    DATA lt_emails TYPE STANDARD TABLE OF ad_smtpadr WITH EMPTY KEY.

    LOOP AT so_mail ASSIGNING FIELD-SYMBOL(<r>).
      IF <r>-low IS NOT INITIAL.
        APPEND CONV ad_smtpadr( <r>-low ) TO lt_emails.
      ENDIF.
    ENDLOOP.
    IF lt_emails IS INITIAL.
      MESSAGE 'Provide recipients in SO_MAIL-LOW (emails)' TYPE 'E'.
    ENDIF.

    DATA lt_body TYPE soli_tab.
    APPEND |Hi,| TO lt_body.
    APPEND |Please find the XLSX attached.| TO lt_body.

    PERFORM send_xlsx_via_email USING lv_xlsx p_subj p_file lt_body lt_emails.
    WRITE: / 'Send requested. Check SOST.'.
  ENDIF.

FORM send_xlsx_via_email
  USING    iv_xlsx     TYPE xstring
           iv_subject  TYPE so_obj_des
           iv_filename TYPE string
           it_body     TYPE soli_tab
           it_emails   TYPE STANDARD TABLE.

  DATA lt_solix TYPE solix_tab.
  DATA lv_size  TYPE sood-objlen.
  DATA lo_doc   TYPE REF TO cl_document_bcs.
  DATA lo_send  TYPE REF TO cl_bcs.
  DATA lt_hdr   TYPE soli_tab.

  DATA lv_fn_str TYPE string.
  DATA lv_fn_so  TYPE so_obj_des.
  DATA lv_hdr    TYPE string.

  lv_fn_str = iv_filename.

  "Avoid CP/AND issues: use ends_with( )
DATA lv_len  TYPE i.
DATA lv_off  TYPE i.
DATA lv_tail TYPE string.

lv_len = strlen( lv_fn_str ).

IF lv_len < 5.
  CONCATENATE lv_fn_str '.xlsx' INTO lv_fn_str.
ELSE.
  lv_off = lv_len - 5.
  lv_tail = lv_fn_str+lv_off(5).

  TRANSLATE lv_tail TO UPPER CASE.

  IF lv_tail <> '.XLSX'.
    CONCATENATE lv_fn_str '.xlsx' INTO lv_fn_str.
  ENDIF.
ENDIF.



  "Attachment subject must be SO_OBJ_DES (50)
  lv_fn_so = CONV so_obj_des( lv_fn_str ).
  IF strlen( lv_fn_str ) > 50.
    lv_fn_so = lv_fn_str+0(50).
  ENDIF.

  CONCATENATE '&SO_FILENAME=' lv_fn_str INTO lv_hdr.
  APPEND lv_hdr TO lt_hdr.

  lt_solix = cl_bcs_convert=>xstring_to_solix( iv_xstring = iv_xlsx ).
  lv_size  = xstrlen( iv_xlsx ).

  lo_doc = cl_document_bcs=>create_document(
             i_type    = 'RAW'
             i_text    = it_body
             i_subject = iv_subject ).

  lo_doc->add_attachment(
    i_attachment_type    = 'XLS'
    i_attachment_subject = lv_fn_so
    i_attachment_size    = lv_size
    i_att_content_hex    = lt_solix
    i_attachment_header  = lt_hdr ).

  lo_send = cl_bcs=>create_persistent( ).
  lo_send->set_document( lo_doc ).

  FIELD-SYMBOLS <e> TYPE any.
  LOOP AT it_emails ASSIGNING <e>.
    DATA(lv_mail) = CONV ad_smtpadr( <e> ).
    lo_send->add_recipient(
      i_recipient = cl_cam_address_bcs=>create_internet_address( lv_mail )
      i_express   = abap_true ).
  ENDLOOP.

  lo_send->send( i_with_error_screen = abap_true ).
  COMMIT WORK.

ENDFORM.