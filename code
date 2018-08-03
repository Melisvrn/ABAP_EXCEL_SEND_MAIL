REPORT ZSB_EXCEL_SEND_MAIL.

TYPES:
      BEGIN OF t_emp_dat,
         tckno       TYPE ZPERS_BILG-tckno     ,
         ad          TYPE ZPERS_BILG-ad        ,
         soyad       TYPE ZPERS_BILG-soyad     ,
         tarih       TYPE ZPERS_BILG-tarih     ,
         avans       TYPE ZPERS_BILG-avans     ,
         kalan_maas  TYPE ZPERS_BILG-kalan_maas,
      END OF t_emp_dat.
DATA:
      w_emp_data TYPE t_emp_dat.
DATA:
      i_emp_data TYPE STANDARD TABLE OF t_emp_dat.
*--------------------------------------------------------*
"  Mail related declarations
*--------------------------------------------------------*
"Variables
DATA :
    g_sent_to_all   TYPE sonv-flag,
    g_tab_lines     TYPE i.
"Types
TYPES:
    t_document_data  TYPE  sodocchgi1,
    t_packing_list   TYPE  sopcklsti1,
    t_attachment     TYPE  solisti1,
    t_body_msg       TYPE  solisti1,
    t_receivers      TYPE  somlreci1.
"Workareas
DATA :
    w_document_data  TYPE  t_document_data,
    w_packing_list   TYPE  t_packing_list,
    w_attachment     TYPE  t_attachment,
    w_body_msg       TYPE  t_body_msg,
    w_receivers      TYPE  t_receivers.
"Internal Tables
DATA :
    i_document_data  TYPE STANDARD TABLE OF t_document_data,
    i_packing_list   TYPE STANDARD TABLE OF t_packing_list,
    i_attachment     TYPE STANDARD TABLE OF t_attachment,
    i_body_msg       TYPE STANDARD TABLE OF t_body_msg,
    i_receivers      TYPE STANDARD TABLE OF t_receivers.

*--------------------------------------------------------*
"Start-of-selection.
*--------------------------------------------------------*
START-OF-SELECTION.
  PERFORM get_data.
  PERFORM build_xls_data_table.

*--------------------------------------------------------*
  "End-of-selection.
*--------------------------------------------------------*
END-OF-SELECTION.
  PERFORM send_mail.

*&--------------------------------------------------------*
  "Form  get_data from ZPERS_BILG
*&--------------------------------------------------------*
FORM get_data.

  SELECT tckno
  ad
  soyad
  tarih
  avans
  kalan_maas
  FROM ZPERS_BILG
  INTO CORRESPONDING FIELDS OF TABLE i_emp_data
  UP TO 4 ROWS.

ENDFORM.                    " get_data
*&---------------------------------------------------------*
"Form  build_xls_data_table
*&---------------------------------------------------------*
FORM build_xls_data_table.
  "If you have Unicode check active in program attributes then
  "you will need to declare constants as follows.
  CLASS cl_abap_char_utilities DEFINITION LOAD.
  CONSTANTS:
  con_tab  TYPE c VALUE cl_abap_char_utilities=>horizontal_tab,
  con_cret TYPE c VALUE cl_abap_char_utilities=>cr_lf.

  CONCATENATE 'TCKNO' 'AD' 'SOYAD' 'TARIH' 'AVANS' 'KALAN_MAAS'
  INTO  w_attachment
  SEPARATED BY  con_tab.

  CONCATENATE con_cret
  w_attachment
  INTO w_attachment.

  APPEND w_attachment TO i_attachment.
  CLEAR  w_attachment.

  LOOP AT i_emp_data INTO w_emp_data.

    CONCATENATE
    w_emp_data-tckno
    w_emp_data-ad
    w_emp_data-soyad
    w_emp_data-tarih
    w_emp_data-avans
    w_emp_data-kalan_maas
    INTO w_attachment
    SEPARATED BY con_tab.

    CONCATENATE con_cret w_attachment
    INTO w_attachment.

    APPEND w_attachment TO i_attachment.
    CLEAR  w_attachment.
  ENDLOOP.

ENDFORM.                    "build_xls_data_table
*&----------------------------------------------------------*
"Form  send_mail
"---------------
"PACKING LIST
"This table requires information about how the data in the
"tables OBJECT_HEADER, CONTENTS_BIN and CONTENTS_TXT are to
"be distributed to the documents and its attachments.The first
"row is for the document, the following rows are each for one
"attachment.
*&-----------------------------------------------------------*
FORM send_mail .

  "Subject of the mail.
  w_document_data-obj_name  = 'MAIL_TO_HEAD'.
  w_document_data-obj_descr = 'Regarding Mail Program by SAP ABAP'.

  "Body of the mail
  PERFORM build_body_of_mail
  USING:space,
  'PERSONEL BILGISI'.

  "Write Packing List for Body
  DESCRIBE TABLE i_body_msg LINES g_tab_lines.
  w_packing_list-head_start = 1.
  w_packing_list-head_num   = 0.
  w_packing_list-body_start = 1.
  w_packing_list-body_num   = g_tab_lines.
  w_packing_list-doc_type   = 'RAW'.
  APPEND w_packing_list TO i_packing_list.
  CLEAR  w_packing_list.

  "Write Packing List for Attachment
  w_packing_list-transf_bin = space.
  w_packing_list-head_start = 1.
  w_packing_list-head_num   = 1.
  w_packing_list-body_start = g_tab_lines + 1.
  DESCRIBE TABLE i_attachment LINES w_packing_list-body_num.
  w_packing_list-doc_type   = 'XLSX'.
  w_packing_list-obj_descr  = 'Excell Attachment'.
  w_packing_list-obj_name   = 'XLSX_ATTACHMENT'.
  w_packing_list-doc_size   = w_packing_list-body_num * 255.
  APPEND w_packing_list TO i_packing_list.
  CLEAR  w_packing_list.

  APPEND LINES OF i_attachment TO i_body_msg.
  "Fill the document data and get size of attachment
  w_document_data-obj_langu  = sy-langu.
  READ TABLE i_body_msg INTO w_body_msg INDEX g_tab_lines.
  w_document_data-doc_size = ( g_tab_lines - 1 ) * 255 + strlen( w_body_msg ).

  "Receivers List.
  w_receivers-rec_type   = 'U'.  "Internet address
  w_receivers-receiver   = 'abc@xyz.com'.
  w_receivers-com_type   = 'INT'.
  w_receivers-notif_del  = 'X'.
  w_receivers-notif_ndel = 'X'.
  APPEND w_receivers TO i_receivers .
  CLEAR:w_receivers.

  "Function module to send mail to Recipients
  CALL FUNCTION 'SO_NEW_DOCUMENT_ATT_SEND_API1'
    EXPORTING
      document_data              = w_document_data
      put_in_outbox              = 'X'
      commit_work                = 'X'
    IMPORTING
      sent_to_all                = g_sent_to_all
    TABLES
      packing_list               = i_packing_list
      contents_txt               = i_body_msg
      receivers                  = i_receivers
    EXCEPTIONS
      too_many_receivers         = 1
      document_not_sent          = 2
      document_type_not_exist    = 3
      operation_no_authorization = 4
      parameter_error            = 5
      x_error                    = 6
      enqueue_error              = 7
      OTHERS                     = 8.

  IF sy-subrc = 0 .
    MESSAGE i303(me) WITH 'Mail has been Successfully Sent.'.
  ELSE.
    WAIT UP TO 2 SECONDS.
    "This program starts the SAPconnect send process.
    SUBMIT rsconn01 WITH mode = 'INT'
    WITH output = 'X'
    AND RETURN.
  ENDIF.

ENDFORM.                    " send_mail
*&-----------------------------------------------------------*
"      Form  build_body_of_mail
*&-----------------------------------------------------------*
FORM build_body_of_mail  USING l_message.

  w_body_msg = l_message.
  APPEND w_body_msg TO i_body_msg.
  CLEAR  w_body_msg.

ENDFORM.                    " build_body_of_mail
