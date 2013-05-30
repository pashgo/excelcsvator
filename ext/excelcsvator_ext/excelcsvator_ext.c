#include <stdlib.h>
#include <stdio.h>
#include <ruby.h>
#include <string.h>
#include <freexl.h>

static char  stringSeparator = '\"';
static char *lineSeparator = "\n";
static char *sheetSeparator = "\f";
static char *fieldSeparator = ",";

static VALUE method_to_c_csv(VALUE self);
static char *print_csv_int(int number);
static char *print_csv_double(double number);
static char *print_csv_string(const char *string);

void Init_excelcsvator_ext(void) {
  VALUE Excelcsvator, ExcelcsvatorExt;
  // , ext_version;

  Excelcsvator = rb_const_get(rb_cObject, rb_intern("Excelcsvator"));
  
  ExcelcsvatorExt = rb_define_module_under(Excelcsvator, "ExcelcsvatorExt");
  // ext_version = rb_str_new2(VERSION);
  // rb_define_const(ExcelcsvatorExt, "VERSION", ext_version);
  rb_define_module_function(ExcelcsvatorExt, "to_c_csv", method_to_c_csv, 0);

  rb_include_module(Excelcsvator, ExcelcsvatorExt);
}

static VALUE method_to_c_csv(VALUE self){
  VALUE buffer = rb_str_new2("");
  VALUE path = rb_funcall(self, rb_intern("path"), 0);

  unsigned int worksheet_index;
  // const char *table_prefix = "xl_table";
  // char table_name[2048];
  const void *handle;
  int ret;
  unsigned int info;
  unsigned int max_worksheet;
  unsigned int rows;
  unsigned short columns;
  unsigned int row;
  unsigned short col;

  /* opening the .XLS file [Workbook] */
  ret = freexl_open(RSTRING_PTR(path), &handle);
  if (ret != FREEXL_OK) {
    // printf("OPEN ERROR: %d\n", ret);
    return buffer;
  }

  /* checking for Password (obfuscated/encrypted) */
  ret = freexl_get_info(handle, FREEXL_BIFF_PASSWORD, &info);
  if (ret != FREEXL_OK) {
    // printf("GET-INFO [FREEXL_BIFF_PASSWORD] Error: %d\n", ret);
    goto stop;
  }

  switch (info) {
    case FREEXL_BIFF_PLAIN:
    break;
    case FREEXL_BIFF_OBFUSCATED:
    default:
      // printf ("Password protected: (not accessible)\n");
      goto stop;
  };

  /* querying BIFF Worksheet entries */
  ret = freexl_get_info (handle, FREEXL_BIFF_SHEET_COUNT, &max_worksheet);
  if (ret != FREEXL_OK) {
    // printf("GET-INFO [FREEXL_BIFF_SHEET_COUNT] Error: %d\n", ret);
    goto stop;
  }

  /* SQL output */
  // printf ("--\n-- this SQL script was automatically created by xl2sql\n");
  // printf ("--\n-- input .xls document was: %s\n--\n", argv[1]);
  // printf ("\nBEGIN;\n\n");

  for (worksheet_index = 0; worksheet_index < max_worksheet; worksheet_index++) {
    const char *utf8_worsheet_name;
    int isFirstLine = 1;
    // make_table_name (table_prefix, worksheet_index, table_name);
    ret = freexl_get_worksheet_name (handle, worksheet_index, &utf8_worsheet_name);
    if (ret != FREEXL_OK) {
      // printf("GET-WORKSHEET-NAME Error: %d\n", ret);
      goto stop;
    }
    /* selecting the active Worksheet */
    ret = freexl_select_active_worksheet (handle, worksheet_index);
    if (ret != FREEXL_OK) {
      // printf("SELECT-ACTIVE_WORKSHEET Error: %d\n", ret);
      goto stop;
    }
    /* dimensions */
    ret = freexl_worksheet_dimensions (handle, &rows, &columns);
    if (ret != FREEXL_OK) {
      // printf("WORKSHEET-DIMENSIONS Error: %d\n", ret);
      goto stop;
    }

    rb_str_append(buffer, rb_str_new2(utf8_worsheet_name));
    rb_str_append(buffer, rb_str_new2(lineSeparator));
    // printf ("--\n-- creating a DB table\n");
    // printf ("-- extracting data from Worksheet #%u: %s\n--\n", worksheet_index, utf8_worsheet_name);
    // printf ("CREATE TABLE %s (\n", table_name);
    // printf ("\trow_no INTEGER NOT NULL PRIMARY KEY");
    // for (col = 0; col < columns; col++)
    //   printf (",\n\tcol_%03u MULTITYPE", col);
    // printf (");\n");

    // printf ("--\n-- populating the same table\n--\n");
    for (row = 0; row < rows; row++) {
      /* INSERT INTO statements */
      FreeXL_CellValue cell;
      int isFirstCol = 1;

      if (!isFirstLine) {
        rb_str_append(buffer, rb_str_new2(lineSeparator));
      } else {
        isFirstLine = 0;
      }
      // printf ("INSERT INTO %s (row_no", table_name);

      // for (col = 0; col < columns; col++)
      //   printf (", col_%03u", col);
      // printf (") VALUES (%u", row);
      for (col = 0; col < columns; col++) {
        ret = freexl_get_cell_value (handle, row, col, &cell);
        if (ret != FREEXL_OK) {
          // fprintf (stderr, "CELL-VALUE-ERROR (r=%u c=%u): %d\n", row, col, ret);
          goto stop;
        }
        
        if (!isFirstCol) {
          rb_str_append(buffer, rb_str_new2(fieldSeparator));
        } else {
          isFirstCol = 0;
        }

        switch (cell.type) {
          case FREEXL_CELL_INT:
            rb_str_append(buffer, rb_str_new2(print_csv_int(cell.value.int_value)));
            // printf (", %d", cell.value.int_value);
            break;
          case FREEXL_CELL_DOUBLE:
            rb_str_append(buffer, rb_str_new2(print_csv_double(cell.value.double_value)));
            // printf (", %1.12f", cell.value.double_value);
            break;
          case FREEXL_CELL_TEXT:
          case FREEXL_CELL_SST_TEXT:
          case FREEXL_CELL_DATE:
          case FREEXL_CELL_DATETIME:
          case FREEXL_CELL_TIME:
            // print_sql_string(cell.value.text_value);
            rb_str_append(buffer, rb_str_new2(print_csv_string(cell.value.text_value)));
            break;
          // case FREEXL_CELL_DATE:
          // case FREEXL_CELL_DATETIME:
          // case FREEXL_CELL_TIME:
          //   rb_str_append(buffer, rb_str_new2(print_csv_string(cell.value.text_value)));
          //   // printf (", '%s'", cell.value.text_value);
          //   break;
          case FREEXL_CELL_NULL:
          default:
            rb_str_append(buffer, rb_str_new2(print_csv_string("")));
            // printf (", NULL");
            break;
        };
      }
      // printf (");\n");
    }

    if (worksheet_index < max_worksheet - 1){
      rb_str_append(buffer, rb_str_new2(lineSeparator));
      rb_str_append(buffer, rb_str_new2(sheetSeparator));
    }
    // printf ("\n-- done: table end\n\n\n\n");
  }
  // printf ("COMMIT;\n");

  stop:
  /* closing the .XLS file [Workbook] */
    ret = freexl_close(handle);
    if (ret != FREEXL_OK) {
      // printf("CLOSE ERROR: %d\n", ret);
      buffer = Qnil;
    }
    return buffer;
}

static char *print_csv_int(int number){
  char *result;
  asprintf(&result, "%d", number);
  return result;
}

static char *print_csv_double(double number){
  char *result;
  asprintf(&result, "%1.12g", number);
  return result;
}

static char *print_csv_string(const char *string) {
  char *result = "";
  const char *p = string;

  asprintf(&result, "%s%c", result, stringSeparator);
  // putchar (stringSeparator);
  while (*p != '\0') {
    if (*p == stringSeparator) {
      asprintf(&result, "%s%c%c", result, stringSeparator, stringSeparator);
    } else if (*p == '\n') {
      asprintf(&result, "%s%c", result, ' ');
    } else if (*p == '\\') {
      asprintf(&result, "%s%s", result, "\\\\");
    } else {
      asprintf(&result, "%s%c", result, *p);
    }
    p++;
  }
  asprintf(&result, "%s%c", result, stringSeparator);
  // putchar (stringSeparator);

  return result;
}