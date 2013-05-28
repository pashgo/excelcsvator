
#include <ruby.h>
#include <xls.h>

static char  stringSeparator = '\"';
static char *lineSeparator = "\n";
static char *sheetSeparator = "\f";
static char *fieldSeparator = ",";
static char *encoding = "UTF-8";

static VALUE buffer;

static void output_string(const char *string, VALUE buffer);
static void output_number(const double number, VALUE buffer);

static VALUE method_to_c_csv(VALUE self){
  xlsWorkBook* pWB;
  xlsWorkSheet* pWS;
  unsigned int i;
  VALUE path;
  struct st_row_data* row;
  DWORD cellRow, cellCol;

  buffer = rb_str_new2("");
  path = rb_funcall(self, rb_intern("path"), 0);

  pWB = xls_open(RSTRING_PTR(path), encoding);
  if (!pWB) {
    return Qnil;
  }

  for (i = 0; i < pWB->sheets.count; i++) {
    int isFirstLine = 1;
    char* sheetName = (char *)pWB->sheets.sheet[i].name;

    rb_str_append(buffer, rb_str_new2(sheetName));
    rb_str_append(buffer, rb_str_new2(lineSeparator));
    
    pWS = xls_getWorkSheet(pWB, i);
    xls_parseWorkSheet(pWS);

    for (cellRow = 0; cellRow <= pWS->rows.lastrow; cellRow++) {
      int isFirstCol = 1;
      row = xls_row(pWS, cellRow);
      // printf("%d form %d\n", cellRow, pWS->rows.lastrow);

      // process cells
      if (!isFirstLine) {
        rb_str_append(buffer, rb_str_new2(lineSeparator));
        // printf("%s", lineSeparator);
      } else {
        isFirstLine = 0;
      }

      for (cellCol = 0; cellCol <= pWS->rows.lastcol; cellCol++) {
        xlsCell *cell = xls_cell(pWS, cellRow, cellCol);

        if ((!cell)) {
          // || (cell->isHidden)
          // verbose("Empty");
          continue;
        }

        if (!isFirstCol) {
          rb_str_append(buffer, rb_str_new2(fieldSeparator));
          // printf("%s", fieldSeparator);
        } else {
          isFirstCol = 0;
        }

        if (cell->rowspan > 1) {
          // fprintf(stderr, "Warning: %d rows spanned at col=%d row=%d: output will not match the Excel file.\n", cell->rowspan, cellCol+1, cellRow+1);
        }

        // display the value of the cell (either numeric or string)
        if (cell->id == 0x27e || cell->id == 0x0BD || cell->id == 0x203) {
          output_number(cell->d, buffer);
        } 
        else if (cell->id == 0x06) {
          // formula
          if (cell->l == 0) // its a number
          {
            output_number(cell->d, buffer);
          } 
          else {
            if (!strcmp((char *)cell->str, "bool")) // its boolean, and test cell->d
            {
              output_string((int) cell->d ? "true" : "false", buffer);
            } else if (!strcmp((char *)cell->str, "error")) // formula is in error
            {
              output_string("*error*", buffer);
            } else // ... cell->str is valid as the result of a string formula.
            {
              output_string((char *)cell->str, buffer);
            }
          }
        } else if (cell->str != NULL) {
          output_string((char *)cell->str, buffer);
        } else {
          output_string("", buffer);
        }
      }
    }

    if (i + 1 < pWB->sheets.count){
      rb_str_append(buffer, rb_str_new2(lineSeparator));
      rb_str_append(buffer, rb_str_new2(sheetSeparator));
    }
    xls_close_WS(pWS);
  }

  xls_close(pWB);

  return buffer;
}

void Init_excelcsvator_ext(void) {
  VALUE Excelcsvator, ExcelcsvatorExt, ext_version;

  Excelcsvator = rb_const_get(rb_cObject, rb_intern("Excelcsvator"));
  
  ExcelcsvatorExt = rb_define_module_under(Excelcsvator, "ExcelcsvatorExt");
  ext_version = rb_str_new2(VERSION);
  rb_define_const(ExcelcsvatorExt, "VERSION", ext_version);
  rb_define_module_function(ExcelcsvatorExt, "to_c_csv", method_to_c_csv, 0);

  rb_include_module(Excelcsvator, ExcelcsvatorExt);
}

static void output_string(const char *string, VALUE buffer) {
  const char *str;
  char *result = "";
  
  asprintf(&result, "%c", stringSeparator);
  rb_str_append(buffer, rb_str_new2(result));
  for (str = string; *str; str++) {
    if (*str == stringSeparator) {
      // asprintf(&result, "%c%c", stringSeparator, stringSeparator);
      asprintf(&result, " ");
      rb_str_append(buffer, rb_str_new2(result));
    } else if (*str == '\\') {
      asprintf(&result, "\\\\");
      rb_str_append(buffer, rb_str_new2(result));
    } else {
      asprintf(&result, "%c", *str);
      rb_str_append(buffer, rb_str_new2(result));
    }
  }
  asprintf(&result, "%c", stringSeparator);
  rb_str_append(buffer, rb_str_new2(result));
  free(result);
}

static void output_number(const double number, VALUE buffer) {
  char *result;
  asprintf(&result, "%.10g", number);
  rb_str_append(buffer, rb_str_new2(result));
  free(result);
}