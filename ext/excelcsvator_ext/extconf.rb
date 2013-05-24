require 'mkmf'

checking_for "-g -O2 -Wall option to compiler" do
  $CFLAGS += " -g -O2 -Wall" if try_compile '', '-g -O2 -Wall'
end

def have_iconv?
  %w{ iconv_open libiconv_open }.any? do |method|
    have_func(method, 'iconv.h') or
      have_library('iconv', method, 'iconv.h') or
      find_library('iconv', method, 'iconv.h')
  end
end

have_func("asprintf")
have_func("malloc")
have_func("realloc")
have_func("strdup")

def have_iconv?
  %w{ iconv_open libiconv_open }.any? do |method|
    have_func(method, 'iconv.h') or
      have_library('iconv', method, 'iconv.h') or
      find_library('iconv', method, 'iconv.h')
  end
end

have_iconv?
have_header("dlfcn.h")
have_header("inttypes.h")
have_header("memory.h")
have_header("stdint.h")
have_header("stdlib.h")
have_header("strings.h")
have_header("string.h")
have_header("sys/stat.h")
have_header("sys/types.h")
have_header("unistd.h")
have_header("wchar.h")

create_makefile('excelcsvator_ext/excelcsvator_ext')