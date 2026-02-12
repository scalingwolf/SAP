CLASS zcl_xlsx_mailer DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    METHODS constructor.

    METHODS add_sheet
      IMPORTING
        iv_title TYPE string
        it_table TYPE ANY TABLE.

    METHODS build_xlsx
      RETURNING VALUE(rv_xlsx) TYPE xstring.

    METHODS send_mail
      IMPORTING
        it_to       TYPE STANDARD TABLE        "table of emails (string/ad_smtpadr)
        iv_subject  TYPE so_obj_des
        it_body     TYPE soli_tab
        iv_filename TYPE string DEFAULT 'report.xlsx'.

  PRIVATE SECTION.

    "===================== strict 7.40 named types =====================
    TYPES ty_c1   TYPE c LENGTH 1.
    TYPES ty_c20  TYPE c LENGTH 20.
    TYPES ty_c80  TYPE c LENGTH 80.
    TYPES ty_c200 TYPE c LENGTH 200.
    TYPES tt_string TYPE STANDARD TABLE OF string WITH DEFAULT KEY.

    "===================== normalized workbook model ===================
    TYPES: BEGIN OF ty_cell,
             kind  TYPE c LENGTH 1,  "T=text, N=number, D=date, M=time
             sval  TYPE string,
             nval  TYPE string,
             style TYPE i,           "cellXfs index
           END OF ty_cell.
    TYPES tt_cells TYPE STANDARD TABLE OF ty_cell WITH DEFAULT KEY.

    TYPES: BEGIN OF ty_row,
             cells TYPE tt_cells,
           END OF ty_row.
    TYPES tt_rows TYPE STANDARD TABLE OF ty_row WITH DEFAULT KEY.

    TYPES: BEGIN OF ty_sheet,
             title   TYPE string,
             headers TYPE tt_string,
             rows    TYPE tt_rows,
           END OF ty_sheet.
    TYPES tt_sheets TYPE STANDARD TABLE OF ty_sheet WITH DEFAULT KEY.

    DATA mt_sheets TYPE tt_sheets.

    "===================== conversions & sanitizers =====================
    METHODS i_to_s
      IMPORTING iv_i TYPE i
      RETURNING VALUE(rv_s) TYPE string.

    METHODS get_user_dcpfm
      RETURNING VALUE(rv_dcpfm) TYPE ty_c1.

    METHODS normalize_numeric_token
      IMPORTING iv_raw TYPE string
      RETURNING VALUE(rv_num) TYPE string.

    METHODS anynum_to_s
      IMPORTING iv_any TYPE any
      RETURNING VALUE(rv_s) TYPE string.

    METHODS ensure_xlsx_ext
      CHANGING cv_filename TYPE string.

    METHODS str_to_xstr_utf8
      IMPORTING iv_text TYPE string
      RETURNING VALUE(rv_xstr) TYPE xstring.

    METHODS sanitize_sheet_name
      IMPORTING iv_title TYPE string
      RETURNING VALUE(rv_title) TYPE string.

    METHODS sanitize_xml_text
      IMPORTING iv_text TYPE string
      RETURNING VALUE(rv_text) TYPE string.

    METHODS esc_xml
      IMPORTING iv_text TYPE string
      RETURNING VALUE(rv_text) TYPE string.

    METHODS esc_xml_text_for_t
      IMPORTING iv_text TYPE string
      RETURNING VALUE(rv_text) TYPE string.

    METHODS col_letters
      IMPORTING iv_idx TYPE i
      RETURNING VALUE(rv_col) TYPE string.

    "===================== excel serials =====================
    METHODS excel_date_serial
      IMPORTING iv_date TYPE d
      RETURNING VALUE(rv) TYPE string.

    METHODS excel_time_serial
      IMPORTING iv_time TYPE t
      RETURNING VALUE(rv) TYPE string.

    "===================== normalize any itab =====================
    METHODS normalize_itab
      IMPORTING it_table   TYPE ANY TABLE
      EXPORTING et_headers TYPE tt_string
                et_rows    TYPE tt_rows.

    "===================== OOXML builders =====================
    METHODS mk_content_types
      IMPORTING iv_sheet_count TYPE i
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS mk_root_rels
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS mk_workbook
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS mk_workbook_rels
      IMPORTING iv_sheet_count TYPE i
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS mk_styles
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS mk_sheet_xml
      IMPORTING is_sheet TYPE ty_sheet
      RETURNING VALUE(rv_xml) TYPE string.

ENDCLASS.



CLASS ZCL_XLSX_MAILER IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_MAILER->ADD_SHEET
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TITLE                       TYPE        STRING
* | [--->] IT_TABLE                       TYPE        ANY TABLE
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD add_sheet.
    DATA ls_sheet TYPE ty_sheet.
    CLEAR ls_sheet.

    IF iv_title IS INITIAL.
      zcx_xlsx_mailer=>raise_msg( 'Sheet title is initial' ).
    ENDIF.

    ls_sheet-title = sanitize_sheet_name( iv_title ).
    IF ls_sheet-title IS INITIAL.
      zcx_xlsx_mailer=>raise_msg( 'Sheet title invalid after sanitizing' ).
    ENDIF.

    normalize_itab(
      EXPORTING
        it_table   = it_table
      IMPORTING
        et_headers = ls_sheet-headers
        et_rows    = ls_sheet-rows ).

    IF ls_sheet-headers IS INITIAL.
      zcx_xlsx_mailer=>raise_msg( 'No columns detected (ITAB line type must be STRUCTURE)' ).
    ENDIF.

    APPEND ls_sheet TO mt_sheets.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->ANYNUM_TO_S
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_ANY                         TYPE        ANY
* | [<-()] RV_S                           TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD anynum_to_s.
    DATA lv_c TYPE ty_c80.
    DATA lv_s TYPE string.

    CLEAR lv_c.
    WRITE iv_any TO lv_c.
    CONDENSE lv_c NO-GAPS.

    lv_s = lv_c.
    rv_s = normalize_numeric_token( lv_s ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_MAILER->BUILD_XLSX
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_XLSX                        TYPE        XSTRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD build_xlsx.
    DATA lo_zip TYPE REF TO cl_abap_zip.
    DATA lv_cnt TYPE i.
    DATA lv_idx TYPE i.

    IF mt_sheets IS INITIAL.
      zcx_xlsx_mailer=>raise_msg( 'No sheets added. Call ADD_SHEET first.' ).
    ENDIF.

    CREATE OBJECT lo_zip.
    lv_cnt = lines( mt_sheets ).

    lo_zip->add( name = '[Content_Types].xml'
                 content = str_to_xstr_utf8( mk_content_types( lv_cnt ) ) ).
    lo_zip->add( name = '_rels/.rels'
                 content = str_to_xstr_utf8( mk_root_rels( ) ) ).
    lo_zip->add( name = 'xl/workbook.xml'
                 content = str_to_xstr_utf8( mk_workbook( ) ) ).
    lo_zip->add( name = 'xl/_rels/workbook.xml.rels'
                 content = str_to_xstr_utf8( mk_workbook_rels( lv_cnt ) ) ).
    lo_zip->add( name = 'xl/styles.xml'
                 content = str_to_xstr_utf8( mk_styles( ) ) ).

    lv_idx = 0.
    LOOP AT mt_sheets INTO DATA(ls_sheet).
      DATA lv_path TYPE string.
      DATA lv_idx_s TYPE string.

      lv_idx = lv_idx + 1.
      lv_idx_s = i_to_s( lv_idx ).
      CONCATENATE 'xl/worksheets/sheet' lv_idx_s '.xml' INTO lv_path.

      lo_zip->add(
        name    = lv_path
        content = str_to_xstr_utf8( mk_sheet_xml( ls_sheet ) ) ).
    ENDLOOP.

    rv_xlsx = lo_zip->save( ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->COL_LETTERS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_IDX                         TYPE        I
* | [<-()] RV_COL                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD col_letters.
    DATA n   TYPE i.
    DATA rem TYPE i.
    DATA ch  TYPE c LENGTH 1.

    n = iv_idx.
    rv_col = ''.

    WHILE n > 0.
      rem = ( n - 1 ) MOD 26.
      ch = sy-abcde+rem(1).
      CONCATENATE ch rv_col INTO rv_col.
      n = ( n - 1 ) DIV 26.
    ENDWHILE.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_MAILER->CONSTRUCTOR
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD constructor.
    CLEAR mt_sheets.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->ENSURE_XLSX_EXT
* +-------------------------------------------------------------------------------------------------+
* | [<-->] CV_FILENAME                    TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD ensure_xlsx_ext.
    DATA lv_len  TYPE i.
    DATA lv_off  TYPE i.
    DATA lv_tail TYPE string.

    lv_len = strlen( cv_filename ).
    IF lv_len < 5.
      CONCATENATE cv_filename '.xlsx' INTO cv_filename.
      RETURN.
    ENDIF.

    lv_off = lv_len - 5.
    lv_tail = cv_filename+lv_off(5).
    TRANSLATE lv_tail TO UPPER CASE.

    IF lv_tail <> '.XLSX'.
      CONCATENATE cv_filename '.xlsx' INTO cv_filename.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->ESC_XML
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TEXT                        TYPE        STRING
* | [<-()] RV_TEXT                        TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD esc_xml.
    DATA lv TYPE string.
    DATA lv_apo TYPE c LENGTH 1.

    lv = sanitize_xml_text( iv_text ).

    REPLACE ALL OCCURRENCES OF '&' IN lv WITH '&amp;'.
    REPLACE ALL OCCURRENCES OF '<' IN lv WITH '&lt;'.
    REPLACE ALL OCCURRENCES OF '>' IN lv WITH '&gt;'.
    REPLACE ALL OCCURRENCES OF '"' IN lv WITH '&quot;'.

    lv_apo = ''''.
    REPLACE ALL OCCURRENCES OF lv_apo IN lv WITH '&apos;'.

    rv_text = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->ESC_XML_TEXT_FOR_T
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TEXT                        TYPE        STRING
* | [<-()] RV_TEXT                        TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD esc_xml_text_for_t.
    "For <t>: preserve spaces; encode newlines safely
    DATA lv TYPE string.
    DATA lv_cr TYPE c LENGTH 1.
    DATA lv_lf TYPE c LENGTH 1.

    lv_cr = cl_abap_char_utilities=>cr_lf+0(1). "CR
    lv_lf = cl_abap_char_utilities=>newline.     "LF

    lv = iv_text.

    "Normalize CRLF/CR to LF
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN lv WITH lv_lf.
    REPLACE ALL OCCURRENCES OF lv_cr IN lv WITH lv_lf.

    "Escape (also sanitizes)
    lv = esc_xml( lv ).

    "Convert LF into entity AFTER escaping
    REPLACE ALL OCCURRENCES OF lv_lf IN lv WITH '&#10;'.

    rv_text = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->EXCEL_DATE_SERIAL
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_DATE                        TYPE        D
* | [<-()] RV                             TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD excel_date_serial.
  DATA base TYPE d VALUE '18991230'.
  DATA days TYPE i.
  days = iv_date - base.
  rv = i_to_s( days ).  "integer only
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->EXCEL_TIME_SERIAL
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TIME                        TYPE        T
* | [<-()] RV                             TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD excel_time_serial.
    DATA secs TYPE i.
    DATA frac TYPE p LENGTH 16 DECIMALS 10.

    secs = ( iv_time+0(2) * 3600 ) + ( iv_time+2(2) * 60 ) + ( iv_time+4(2) ).
    frac = secs / 86400.

    rv = anynum_to_s( frac ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->GET_USER_DCPFM
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_DCPFM                       TYPE        TY_C1
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_user_dcpfm.
    DATA lv TYPE ty_c1.
    CLEAR lv.
    GET PARAMETER ID 'DCPFM' FIELD lv.
    IF lv IS INITIAL.
      lv = ' '. "default
    ENDIF.
    rv_dcpfm = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->I_TO_S
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_I                           TYPE        I
* | [<-()] RV_S                           TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD i_to_s.
    DATA lv_c TYPE ty_c20.
    CLEAR lv_c.
    WRITE iv_i TO lv_c.
    CONDENSE lv_c NO-GAPS.
    REPLACE ALL OCCURRENCES OF '.' IN lv_c WITH ''.
    REPLACE ALL OCCURRENCES OF ',' IN lv_c WITH ''.
    rv_s = lv_c.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->MK_CONTENT_TYPES
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_SHEET_COUNT                 TYPE        I
* | [<-()] RV_XML                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD mk_content_types.
    DATA lv TYPE string.
    DATA i TYPE i.

    lv =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' &&
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' &&
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' &&
      '<Default Extension="xml" ContentType="application/xml"/>' &&
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' &&
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'.

    DO iv_sheet_count TIMES.
      i = sy-index.
      lv = lv &&
        '<Override PartName="/xl/worksheets/sheet' && i_to_s( i ) &&
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
    ENDDO.

    lv = lv && '</Types>'.
    rv_xml = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->MK_ROOT_RELS
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_XML                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD mk_root_rels.
    rv_xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' &&
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' &&
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' &&
      '</Relationships>'.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->MK_SHEET_XML
* +-------------------------------------------------------------------------------------------------+
* | [--->] IS_SHEET                       TYPE        TY_SHEET
* | [<-()] RV_XML                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD mk_sheet_xml.
    DATA lv    TYPE string.
    DATA col_i TYPE i.
    DATA row_i TYPE i.

    "Compute dimension: A1 : lastcol lastrow (include header row)
    DATA lv_last_col TYPE i.
    DATA lv_last_row TYPE i.
    lv_last_col = lines( is_sheet-headers ).
    lv_last_row = lines( is_sheet-rows ) + 1. "header + data
    IF lv_last_col <= 0.
      lv_last_col = 1.
    ENDIF.
    IF lv_last_row <= 1.
      lv_last_row = 1.
    ENDIF.

    DATA lv_dim TYPE string.
    lv_dim = 'A1:' && col_letters( lv_last_col ) && i_to_s( lv_last_row ).

    lv =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' &&
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' &&
      '<dimension ref="' && lv_dim && '"/>' &&
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' &&
      '<sheetData>'.

    "Header row r=1
    lv = lv && '<row r="1">'.
    col_i = 0.
    LOOP AT is_sheet-headers INTO DATA(h).
      DATA refh TYPE string.
      col_i = col_i + 1.
      refh = col_letters( col_i ).
      CONCATENATE refh '1' INTO refh.

      lv = lv &&
        '<c r="' && refh &&
        '" t="inlineStr" s="1"><is><t xml:space="preserve">' &&
        esc_xml_text_for_t( h ) &&
        '</t></is></c>'.
    ENDLOOP.
    lv = lv && '</row>'.

    "Data rows
    row_i = 1.
    LOOP AT is_sheet-rows INTO DATA(ls_row).
      DATA row_s TYPE string.
      row_i = row_i + 1.
      row_s = i_to_s( row_i ).

      lv = lv && '<row r="' && row_s && '">'.

      col_i = 0.
      LOOP AT ls_row-cells INTO DATA(cel).
        DATA ref TYPE string.
        col_i = col_i + 1.

        ref = col_letters( col_i ).
        CONCATENATE ref row_s INTO ref.

        IF cel-kind = 'N' OR cel-kind = 'D' OR cel-kind = 'M'.

          IF cel-nval IS INITIAL.
            CONTINUE.
          ENDIF.

          DATA lv_num TYPE string.
          lv_num = normalize_numeric_token( cel-nval ).
          IF lv_num IS INITIAL.
            CONTINUE.
          ENDIF.

          "Validate allowed chars for Excel numeric token
          DATA lv_bad TYPE abap_bool.
          DATA lv_len TYPE i.
          DATA lv_pos TYPE i.
          DATA lv_ch  TYPE c LENGTH 1.

          lv_bad = abap_false.
          lv_len = strlen( lv_num ).

          DO lv_len TIMES.
            lv_pos = sy-index - 1.
            lv_ch  = lv_num+lv_pos(1).
            IF ( lv_ch >= '0' AND lv_ch <= '9' ) OR
               lv_ch = '.' OR lv_ch = '-' OR
               lv_ch = 'E' OR lv_ch = 'e'.
              "ok
            ELSE.
              lv_bad = abap_true.
              EXIT.
            ENDIF.
          ENDDO.

          IF lv_bad = abap_true.
            "Fallback to text (never break the sheet)
            lv = lv &&
              '<c r="' && ref && '" t="inlineStr"><is><t xml:space="preserve">' &&
              esc_xml_text_for_t( cel-nval ) &&
              '</t></is></c>'.
            CONTINUE.
          ENDIF.

          DATA lv_style_i TYPE i.
          DATA lv_style_s TYPE string.

          lv_style_i = cel-style.

          IF cel-kind = 'D'.
            lv_style_i = 2. "force date format
          ELSEIF cel-kind = 'M'.
            lv_style_i = 4. "force time format
          ENDIF.

          IF lv_style_i > 0.
            lv_style_s = i_to_s( lv_style_i ).
            lv = lv &&
              '<c r="' && ref && '" s="' && lv_style_s &&
              '"><v>' && lv_num && '</v></c>'.
          ELSE.
            lv = lv &&
              '<c r="' && ref && '"><v>' && lv_num && '</v></c>'.
          ENDIF.

        ELSE.
          lv = lv &&
            '<c r="' && ref && '" t="inlineStr"><is><t xml:space="preserve">' &&
            esc_xml_text_for_t( cel-sval ) &&
            '</t></is></c>'.
        ENDIF.
      ENDLOOP.

      lv = lv && '</row>'.
    ENDLOOP.

    lv = lv && '</sheetData></worksheet>'.
    rv_xml = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->MK_STYLES
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_XML                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD mk_styles.
    "cellXfs index usage:
    "0 general
    "1 header bold
    "2 date yyyy-mm-dd
    "3 spare
    "4 time hh:mm:ss
    rv_xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' &&
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' &&

      '<numFmts count="2">' &&
      '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>' &&
      '<numFmt numFmtId="165" formatCode="hh:mm:ss"/>' &&
      '</numFmts>' &&

      '<fonts count="2">' &&
      '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' &&
      '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' &&
      '</fonts>' &&

      '<fills count="2">' &&
      '<fill><patternFill patternType="none"/></fill>' &&
      '<fill><patternFill patternType="gray125"/></fill>' &&
      '</fills>' &&

      '<borders count="1">' &&
      '<border><left/><right/><top/><bottom/><diagonal/></border>' &&
      '</borders>' &&

      '<cellStyleXfs count="1">' &&
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' &&
      '</cellStyleXfs>' &&

      '<cellXfs count="5">' &&
      '<xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>' &&
      '<xf numFmtId="0"   fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' &&
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' &&
      '<xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>' &&
      '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' &&
      '</cellXfs>' &&

      '<cellStyles count="1">' &&
      '<cellStyle name="Normal" xfId="0" builtinId="0"/>' &&
      '</cellStyles>' &&

      '<dxfs count="0"/>' &&
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' &&

      '</styleSheet>'.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->MK_WORKBOOK
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_XML                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD mk_workbook.
    DATA lv TYPE string.
    DATA idx TYPE i.

    lv =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' &&
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' &&
      '<sheets>'.

    idx = 0.
    LOOP AT mt_sheets INTO DATA(ls).
      DATA idx_s TYPE string.
      idx = idx + 1.
      idx_s = i_to_s( idx ).

      lv = lv &&
        '<sheet name="' && esc_xml( ls-title ) &&
        '" sheetId="' && idx_s &&
        '" r:id="rId' && idx_s && '"/>'.
    ENDLOOP.

    lv = lv && '</sheets></workbook>'.
    rv_xml = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->MK_WORKBOOK_RELS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_SHEET_COUNT                 TYPE        I
* | [<-()] RV_XML                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD mk_workbook_rels.
    DATA lv TYPE string.
    DATA i TYPE i.
    DATA last_i TYPE i.

    lv =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' &&
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.

    DO iv_sheet_count TIMES.
      i = sy-index.
      lv = lv &&
        '<Relationship Id="rId' && i_to_s( i ) &&
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' &&
        i_to_s( i ) && '.xml"/>'.
    ENDDO.

    last_i = iv_sheet_count + 1.
    lv = lv &&
      '<Relationship Id="rId' && i_to_s( last_i ) &&
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' &&
      '</Relationships>'.

    rv_xml = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->NORMALIZE_ITAB
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_TABLE                       TYPE        ANY TABLE
* | [<---] ET_HEADERS                     TYPE        TT_STRING
* | [<---] ET_ROWS                        TYPE        TT_ROWS
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD normalize_itab.
    FIELD-SYMBOLS: <ls>  TYPE any,
                   <val> TYPE any.

    DATA lo_tdesc TYPE REF TO cl_abap_tabledescr.
    DATA lo_line  TYPE REF TO cl_abap_typedescr.
    DATA lo_str   TYPE REF TO cl_abap_structdescr.
    DATA lt_comp  TYPE abap_component_tab.

    TRY.
        lo_tdesc ?= cl_abap_typedescr=>describe_by_data( it_table ).
        lo_line  = lo_tdesc->get_table_line_type( ).
        lo_str ?= lo_line.
      CATCH cx_root.
        zcx_xlsx_mailer=>raise_msg( 'ITAB line type must be a STRUCTURE' ).
    ENDTRY.

    lt_comp = lo_str->get_components( ).

    CLEAR et_headers.
    LOOP AT lt_comp INTO DATA(ls_comp).
      APPEND ls_comp-name TO et_headers.
    ENDLOOP.

    CLEAR et_rows.

    LOOP AT it_table ASSIGNING <ls>.
      DATA ls_outrow TYPE ty_row.
      CLEAR ls_outrow-cells.

      LOOP AT lt_comp INTO ls_comp.
        DATA ls_cell TYPE ty_cell.
        DATA lo_type TYPE REF TO cl_abap_typedescr.
        DATA lv_kind TYPE c LENGTH 1.

        CLEAR ls_cell.

        ASSIGN COMPONENT ls_comp-name OF STRUCTURE <ls> TO <val>.
        IF sy-subrc <> 0.
          ls_cell-kind  = 'T'.
          ls_cell-sval  = ``.
          ls_cell-style = 0.
          APPEND ls_cell TO ls_outrow-cells.
          CONTINUE.
        ENDIF.

        lo_type = ls_comp-type.
        lv_kind = lo_type->type_kind.

        "Default TEXT
        ls_cell-kind  = 'T'.
        ls_cell-style = 0.
        ls_cell-sval  = |{ <val> }|.

        "NUMERIC: I(4), 8(8byte), P(packed), F(float)
        IF lv_kind = 'I' OR lv_kind = '8' OR lv_kind = 'P' OR lv_kind = 'F'.
          ls_cell-kind  = 'N'.
          ls_cell-style = 0.
          ls_cell-nval  = anynum_to_s( <val> ).
          APPEND ls_cell TO ls_outrow-cells.
          CONTINUE.
        ENDIF.

        "DATE: DATS
        IF lv_kind = 'D'.
          DATA lv_d TYPE d.
          lv_d = <val>.
          IF lv_d IS INITIAL.
            ls_cell-kind  = 'T'.
            ls_cell-sval  = ``.
            ls_cell-style = 0.
          ELSE.
            ls_cell-kind  = 'D'.
            ls_cell-style = 2.
            DATA lv_days TYPE string.
lv_days = excel_date_serial( lv_d ).
"force integer: remove anything after decimal if present
SPLIT lv_days AT '.' INTO lv_days DATA(lv_dummy).
ls_cell-nval = lv_days.

          ENDIF.
          APPEND ls_cell TO ls_outrow-cells.
          CONTINUE.
        ENDIF.

        "TIME: TIMS
        IF lv_kind = 'T'.
          DATA lv_t TYPE t.
          lv_t = <val>.
          IF lv_t IS INITIAL.
            ls_cell-kind  = 'T'.
            ls_cell-sval  = ``.
            ls_cell-style = 0.
          ELSE.
            ls_cell-kind  = 'M'.
            ls_cell-style = 4.
            ls_cell-nval  = excel_time_serial( lv_t ).
          ENDIF.
          APPEND ls_cell TO ls_outrow-cells.
          CONTINUE.
        ENDIF.

        "Fallback TEXT
        APPEND ls_cell TO ls_outrow-cells.
      ENDLOOP.

      APPEND ls_outrow TO et_rows.
    ENDLOOP.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->NORMALIZE_NUMERIC_TOKEN
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_RAW                         TYPE        STRING
* | [<-()] RV_NUM                         TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD normalize_numeric_token.
    "Goal: strict Excel numeric token:
    "  - only digits, one leading '-', one '.', optional exponent E/e with optional '-'
    "Input may contain thousand separators and comma decimal based on user format.
    DATA lv TYPE string.
    DATA lv_pos_dot TYPE i.
    DATA lv_pos_com TYPE i.
    DATA lv_len TYPE i.
    DATA lv_i TYPE i.
    DATA lv_ch TYPE c LENGTH 1.

    lv = iv_raw.
    CONDENSE lv NO-GAPS.

    IF lv IS INITIAL.
      rv_num = ''.
      RETURN.
    ENDIF.

    "Remove leading '+'
    IF lv+0(1) = '+'.
      lv = lv+1.
    ENDIF.

    "If both '.' and ',' exist -> decide decimal by last occurrence
    lv_pos_dot = -1.
    lv_pos_com = -1.
    lv_len = strlen( lv ).

    DO lv_len TIMES.
      lv_i = sy-index - 1.
      lv_ch = lv+lv_i(1).
      IF lv_ch = '.'.
        lv_pos_dot = lv_i.
      ELSEIF lv_ch = ','.
        lv_pos_com = lv_i.
      ENDIF.
    ENDDO.

    IF lv_pos_dot >= 0 AND lv_pos_com >= 0.
      "Decimal sep is the one that appears last; other is thousand sep
      IF lv_pos_dot > lv_pos_com.
        "dot is decimal => remove commas
        REPLACE ALL OCCURRENCES OF ',' IN lv WITH ''.
      ELSE.
        "comma is decimal => remove dots then change comma to dot
        REPLACE ALL OCCURRENCES OF '.' IN lv WITH ''.
        REPLACE ALL OCCURRENCES OF ',' IN lv WITH '.'.
      ENDIF.
    ELSEIF lv_pos_com >= 0 AND lv_pos_dot < 0.
      "only comma exists -> treat as decimal comma
      REPLACE ALL OCCURRENCES OF ',' IN lv WITH '.'.
    ELSE.
      "only dot or neither -> ok
    ENDIF.

    CONDENSE lv NO-GAPS.
    rv_num = lv.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->SANITIZE_SHEET_NAME
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TITLE                       TYPE        STRING
* | [<-()] RV_TITLE                       TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD sanitize_sheet_name.
    rv_title = iv_title.

    REPLACE ALL OCCURRENCES OF ':' IN rv_title WITH ' '.
    REPLACE ALL OCCURRENCES OF '\' IN rv_title WITH ' '.
    REPLACE ALL OCCURRENCES OF '/' IN rv_title WITH ' '.
    REPLACE ALL OCCURRENCES OF '?' IN rv_title WITH ' '.
    REPLACE ALL OCCURRENCES OF '*' IN rv_title WITH ' '.
    REPLACE ALL OCCURRENCES OF '[' IN rv_title WITH ' '.
    REPLACE ALL OCCURRENCES OF ']' IN rv_title WITH ' '.

    CONDENSE rv_title.

    IF strlen( rv_title ) > 31.
      rv_title = rv_title+0(31).
      CONDENSE rv_title.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->SANITIZE_XML_TEXT
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TEXT                        TYPE        STRING
* | [<-()] RV_TEXT                        TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD sanitize_xml_text.
    "Remove illegal XML 1.0 control chars except TAB(09), LF(0A), CR(0D)
    DATA lv_in  TYPE string.
    DATA lv_out TYPE string.
    DATA lv_len TYPE i.
    DATA lv_i   TYPE i.
    DATA lv_ch  TYPE c LENGTH 1.

    lv_in = iv_text.
    lv_out = ''.
    lv_len = strlen( lv_in ).

    DO lv_len TIMES.
      lv_i = sy-index - 1.
      lv_ch = lv_in+lv_i(1).

      IF lv_ch = cl_abap_char_utilities=>horizontal_tab OR
         lv_ch = cl_abap_char_utilities=>newline OR
         lv_ch = cl_abap_char_utilities=>cr_lf+0(1) OR
         lv_ch >= ' '.
        lv_out = lv_out && lv_ch.
      ENDIF.
    ENDDO.

    rv_text = lv_out.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_XLSX_MAILER->SEND_MAIL
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_TO                          TYPE        STANDARD TABLE
* | [--->] IV_SUBJECT                     TYPE        SO_OBJ_DES
* | [--->] IT_BODY                        TYPE        SOLI_TAB
* | [--->] IV_FILENAME                    TYPE        STRING (default ='report.xlsx')
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD send_mail.
    DATA lv_xlsx TYPE xstring.
    DATA lv_fn   TYPE string.

    lv_xlsx = build_xlsx( ).

    lv_fn = iv_filename.
    ensure_xlsx_ext( CHANGING cv_filename = lv_fn ).

    DATA lv_att_subj TYPE so_obj_des.
    lv_att_subj = lv_fn.
    IF strlen( lv_fn ) > 50.
      lv_att_subj = lv_fn+0(50).
    ENDIF.

    DATA lt_solix TYPE solix_tab.
    DATA lv_size  TYPE sood-objlen.
    DATA lt_hdr   TYPE soli_tab.
    DATA lv_hdr   TYPE string.

    lt_solix = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xlsx ).
    lv_size  = xstrlen( lv_xlsx ).

    CONCATENATE '&SO_FILENAME=' lv_fn INTO lv_hdr.
    APPEND lv_hdr TO lt_hdr.

    DATA lo_doc  TYPE REF TO cl_document_bcs.
    DATA lo_send TYPE REF TO cl_bcs.

    TRY.
        lo_doc = cl_document_bcs=>create_document(
                   i_type    = 'RAW'
                   i_text    = it_body
                   i_subject = iv_subject ).

        lo_doc->add_attachment(
          i_attachment_type    = 'XLS'
          i_attachment_subject = lv_att_subj
          i_attachment_size    = lv_size
          i_att_content_hex    = lt_solix
          i_attachment_header  = lt_hdr ).

        lo_send = cl_bcs=>create_persistent( ).
        lo_send->set_document( lo_doc ).

        FIELD-SYMBOLS <to> TYPE any.
        LOOP AT it_to ASSIGNING <to>.
          DATA lv_mail TYPE ad_smtpadr.
          CLEAR lv_mail.
          lv_mail = <to>.
          IF lv_mail IS INITIAL.
            CONTINUE.
          ENDIF.

          lo_send->add_recipient(
            i_recipient = cl_cam_address_bcs=>create_internet_address( lv_mail )
            i_express   = abap_true ).
        ENDLOOP.

        lo_send->send( i_with_error_screen = abap_true ).
        COMMIT WORK.

      CATCH cx_bcs INTO DATA(lx_bcs).
        ROLLBACK WORK.
        zcx_xlsx_mailer=>raise_msg( lx_bcs->get_text( ) ).
    ENDTRY.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_XLSX_MAILER->STR_TO_XSTR_UTF8
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TEXT                        TYPE        STRING
* | [<-()] RV_XSTR                        TYPE        XSTRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD str_to_xstr_utf8.
    "Critical: OOXML declares UTF-8, so bytes must be UTF-8
    DATA lo_conv TYPE REF TO cl_abap_conv_out_ce.

    TRY.
        lo_conv = cl_abap_conv_out_ce=>create( encoding = 'UTF-8' ).
        lo_conv->write( data = iv_text ).
        rv_xstr = lo_conv->get_buffer( ).
      CATCH cx_root INTO DATA(lx).
        zcx_xlsx_mailer=>raise_msg( lx->get_text( ) ).
    ENDTRY.
  ENDMETHOD.
ENDCLASS.