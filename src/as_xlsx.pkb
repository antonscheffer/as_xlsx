create or replace package body as_xlsx
is
  --
  c_version constant varchar2(20) := 'as_xlsx47';
--
  c_lob_duration constant pls_integer := dbms_lob.call;
  c_LOCAL_FILE_HEADER        constant raw(4) := hextoraw( '504B0304' ); -- Local file header signature
  c_CENTRAL_FILE_HEADER      constant raw(4) := hextoraw( '504B0102' ); -- Central directory file header signature
  c_END_OF_CENTRAL_DIRECTORY constant raw(4) := hextoraw( '504B0506' ); -- End of central directory signature
--
  type tp_XF_fmt is record
    ( numFmtId pls_integer
    , fontId pls_integer
    , fillId pls_integer
    , borderId pls_integer
    , alignment tp_alignment
    , height number
    );
  type tp_col_fmts is table of tp_XF_fmt index by pls_integer;
  type tp_row_fmts is table of tp_XF_fmt index by pls_integer;
  type tp_widths is table of number index by pls_integer;
  type tp_cell is record
    ( value number
    , style varchar2(50)
    );
  type tp_cells is table of tp_cell index by pls_integer;
  type tp_rows is table of tp_cells index by pls_integer;
  type tp_autofilter is record
    ( column_start pls_integer
    , column_end   pls_integer
    , row_start    pls_integer
    , row_end      pls_integer
    );
  type tp_autofilters is table of tp_autofilter index by pls_integer;
  type tp_table is record
    ( sheet        pls_integer
    , column_start pls_integer
    , column_end   pls_integer
    , row_start    pls_integer
    , row_end      pls_integer
    , style        varchar2(1000)
    , name         varchar2(32767)
    );
  type tp_tables is table of tp_table index by pls_integer;
  type tp_hyperlink is record
    ( cell varchar2(10)
    , url  varchar2(1000)
    , location varchar2(100)
    , tooltip varchar2(1000)
    );
  type tp_hyperlinks is table of tp_hyperlink index by pls_integer;
  subtype tp_author is varchar2(32767 char);
  type tp_authors is table of pls_integer index by tp_author;
  authors tp_authors;
  type tp_comment is record
    ( text varchar2(32767 char)
    , author tp_author
    , row pls_integer
    , column pls_integer
    , width pls_integer
    , height pls_integer
    );
  type tp_comments is table of tp_comment index by pls_integer;
  type tp_mergecells is table of varchar2(21) index by pls_integer;
  type tp_validation is record
    ( type varchar2(10)
    , errorstyle varchar2(32)
    , showinputmessage boolean
    , prompt varchar2(32767 char)
    , title varchar2(32767 char)
    , error_title varchar2(32767 char)
    , error_txt varchar2(32767 char)
    , showerrormessage boolean
    , formula1 varchar2(32767 char)
    , formula2 varchar2(32767 char)
    , allowBlank boolean
    , sqref varchar2(32767 char)
    );
  type tp_validations is table of tp_validation index by pls_integer;
  type tp_drawing is record
    ( img_id pls_integer
    , row pls_integer
    , col pls_integer
    , scale number
    , name varchar2(100)
    , title varchar2(100)
    , description varchar2(4000)
    );
  type tp_drawings is table of tp_drawing index by pls_integer;
  type tp_sheet is record
    ( rows tp_rows
    , widths tp_widths
    , name varchar2(100)
    , freeze_rows pls_integer
    , freeze_cols pls_integer
    , autofilters tp_autofilters
    , hyperlinks tp_hyperlinks
    , col_fmts tp_col_fmts
    , row_fmts tp_row_fmts
    , comments tp_comments
    , mergecells tp_mergecells
    , validations tp_validations
    , drawings tp_drawings
    , tabcolor varchar2(8)
    , show_gridlines boolean
    , grid_color_idx pls_integer
    , show_headers   boolean
    );
  type tp_sheets is table of tp_sheet index by pls_integer;
  type tp_numFmt is record
    ( numFmtId pls_integer
    , formatCode varchar2(100)
    );
  type tp_numFmts is table of tp_numFmt index by pls_integer;
  type tp_fill is record
    ( patternType varchar2(30)
    , fgRGB varchar2(8)
    );
  type tp_fills is table of tp_fill index by pls_integer;
  type tp_cellXfs is table of tp_xf_fmt index by pls_integer;
  type tp_font is record
    ( name varchar2(100)
    , family pls_integer
    , fontsize number
    , theme pls_integer
    , RGB varchar2(8)
    , underline boolean
    , italic boolean
    , bold boolean
    );
  type tp_fonts is table of tp_font index by pls_integer;
  type tp_border is record
    ( top    varchar2(17)
    , bottom varchar2(17)
    , left   varchar2(17)
    , right  varchar2(17)
    , rgb    varchar2(8)
    );
  type tp_borders is table of tp_border index by pls_integer;
  type tp_numFmtIndexes is table of pls_integer index by pls_integer;
  type tp_strings is table of pls_integer index by varchar2(32767 char);
  type tp_str_ind is table of varchar2(32767 char) index by pls_integer;
  type tp_defined_name is record
    ( name varchar2(32767 char)
    , ref varchar2(32767 char)
    , sheet pls_integer
    );
  type tp_defined_names is table of tp_defined_name index by pls_integer;
  type tp_image is record
    ( img blob
    , hash raw(4)
    , width  pls_integer
    , height pls_integer
    );
  type tp_images is table of tp_image index by pls_integer;
  type tp_book is record
    ( sheets tp_sheets
    , strings tp_strings
    , str_ind tp_str_ind
    , str_cnt pls_integer := 0
    , fonts tp_fonts
    , fills tp_fills
    , borders tp_borders
    , numFmts tp_numFmts
    , cellXfs tp_cellXfs
    , numFmtIndexes tp_numFmtIndexes
    , defined_names tp_defined_names
    , images        tp_images
    , tables        tp_tables
    );
  workbook tp_book;
  --
  type tp_zip_info is record
    ( len integer
    , cnt integer
    , len_cd integer
    , idx_cd integer
    , idx_eocd integer
    );
  type tp_cfh is record
    ( offset integer
    , compressed_len integer
    , original_len integer
    , len pls_integer
    , n   pls_integer
    , m   pls_integer
    , k   pls_integer
    , utf8 boolean
    , encrypted boolean
    , crc32 raw(4)
    , external_file_attr raw(4)
    , encoding varchar2(3999)
    , idx   integer
    , name1 raw(32767)
    );
  --
  g_useXf boolean := true;
  --
  g_addtxt2utf8blob_tmp varchar2(32767);
  procedure addtxt2utf8blob_init( p_blob in out nocopy blob )
  is
  begin
    g_addtxt2utf8blob_tmp := null;
    dbms_lob.createtemporary( p_blob, true );
  end;
  procedure addtxt2utf8blob_finish( p_blob in out nocopy blob )
  is
    t_raw raw(32767);
  begin
    t_raw := utl_i18n.string_to_raw( g_addtxt2utf8blob_tmp, 'AL32UTF8' );
    dbms_lob.writeappend( p_blob, utl_raw.length( t_raw ), t_raw );
  exception
    when value_error
    then
      t_raw := utl_i18n.string_to_raw( substr( g_addtxt2utf8blob_tmp, 1, 16381 ), 'AL32UTF8' );
      dbms_lob.writeappend( p_blob, utl_raw.length( t_raw ), t_raw );
      t_raw := utl_i18n.string_to_raw( substr( g_addtxt2utf8blob_tmp, 16382 ), 'AL32UTF8' );
      dbms_lob.writeappend( p_blob, utl_raw.length( t_raw ), t_raw );
  end;
  procedure addtxt2utf8blob( p_txt varchar2, p_blob in out nocopy blob )
  is
  begin
    g_addtxt2utf8blob_tmp := g_addtxt2utf8blob_tmp || p_txt;
  exception
    when value_error
    then
      addtxt2utf8blob_finish( p_blob );
      g_addtxt2utf8blob_tmp := p_txt;
  end;
--
  procedure blob2file
    ( p_blob blob
    , p_directory varchar2 := 'MY_DIR'
    , p_filename varchar2 := 'my.xlsx'
    )
  is
$IF as_xlsx.use_utl_file
$THEN
    t_fh utl_file.file_type;
    t_len pls_integer := 32767;
  begin
    t_fh := utl_file.fopen( p_directory
                          , p_filename
                          , 'wb'
                          );
    for i in 0 .. trunc( ( dbms_lob.getlength( p_blob ) - 1 ) / t_len )
    loop
      utl_file.put_raw( t_fh
                      , dbms_lob.substr( p_blob
                                       , t_len
                                       , i * t_len + 1
                                       )
                      );
    end loop;
    utl_file.fclose( t_fh );
$ELSE
  begin
    raise_application_error( -20024, 'utl_file not available. Change the package header, set as_xlsx.use_utl_file := true; when you have access to utl_file.' );
$END
  end;
--
  function raw2num( p_raw raw, p_len integer, p_pos integer )
  return number
  is
  begin
    return utl_raw.cast_to_binary_integer( utl_raw.substr( p_raw, p_pos, p_len ), utl_raw.little_endian );
  end;
--
  function little_endian( p_big number, p_bytes pls_integer := 4 )
  return raw
  is
  begin
    if p_big < 0
    then
      return utl_raw.reverse( to_char( 4294967296 + p_big, 'fm0XXXXXXX' ) );
    else
      return utl_raw.reverse( to_char( p_big, substr( 'fm0XXXXXXXXXXXXXXXXXXX', 1, 2 + 2 * p_bytes ) ) );
    end if;
  end;
  --
  function little_endian( p_num raw, p_pos pls_integer := 1, p_bytes pls_integer := null )
  return integer
  is
  begin
    return to_number( utl_raw.reverse( utl_raw.substr( p_num, p_pos, p_bytes ) ), 'XXXXXXXXXXXXXXXX' );
  end;
  --
  function blob2num( p_blob blob, p_len integer, p_pos integer )
  return number
  is
  begin
    return utl_raw.cast_to_binary_integer( dbms_lob.substr( p_blob, p_len, p_pos ), utl_raw.little_endian );
  end;
--
  procedure add1file
    ( p_zipped_blob in out blob
    , p_name varchar2
    , p_content blob
    )
  is
    t_now date;
    t_blob blob;
    t_len integer;
    t_clen integer;
    t_crc32 raw(4) := hextoraw( '00000000' );
    t_compressed boolean := false;
    t_name raw(32767);
  begin
    t_now := sysdate;
    t_len := nvl( dbms_lob.getlength( p_content ), 0 );
    if t_len > 0
    then
      t_blob := utl_compress.lz_compress( p_content );
      t_clen := dbms_lob.getlength( t_blob ) - 18;
      t_compressed := t_clen < t_len;
      t_crc32 := dbms_lob.substr( t_blob, 4, t_clen + 11 );
    end if;
    if not t_compressed
    then
      t_clen := t_len;
      t_blob := p_content;
    end if;
    if p_zipped_blob is null
    then
      dbms_lob.createtemporary( p_zipped_blob, true );
    end if;
    t_name := utl_i18n.string_to_raw( p_name, 'AL32UTF8' );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_LOCAL_FILE_HEADER -- Local file header signature
                                   , hextoraw( '1400' )  -- version 2.0
                                   , case when t_name = utl_i18n.string_to_raw( p_name, 'US8PC437' )
                                       then hextoraw( '0000' ) -- no General purpose bits
                                       else hextoraw( '0008' ) -- set Language encoding flag (EFS)
                                     end
                                   , case when t_compressed
                                        then hextoraw( '0800' ) -- deflate
                                        else hextoraw( '0000' ) -- stored
                                     end
                                   , little_endian( to_number( to_char( t_now, 'ss' ) ) / 2
                                                  + to_number( to_char( t_now, 'mi' ) ) * 32
                                                  + to_number( to_char( t_now, 'hh24' ) ) * 2048
                                                  , 2
                                                  ) -- File last modification time
                                   , little_endian( to_number( to_char( t_now, 'dd' ) )
                                                  + to_number( to_char( t_now, 'mm' ) ) * 32
                                                  + ( to_number( to_char( t_now, 'yyyy' ) ) - 1980 ) * 512
                                                  , 2
                                                  ) -- File last modification date
                                   , t_crc32 -- CRC-32
                                   , little_endian( t_clen )                      -- compressed size
                                   , little_endian( t_len )                       -- uncompressed size
                                   , little_endian( utl_raw.length( t_name ), 2 ) -- File name length
                                   , hextoraw( '0000' )                           -- Extra field length
                                   , t_name                                       -- File name
                                   )
                   );
    if t_compressed
    then
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 11 ); -- compressed content
    elsif t_clen > 0
    then
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 1 ); --  content
    end if;
    if dbms_lob.istemporary( t_blob ) = 1
    then
      dbms_lob.freetemporary( t_blob );
    end if;
  end;
--
  procedure finish_zip( p_zipped_blob in out blob )
  is
    t_cnt pls_integer := 0;
    t_offs integer;
    t_offs_dir_header integer;
    t_offs_end_header integer;
    t_comment raw(32767) := utl_raw.cast_to_raw( 'Implementation by Anton Scheffer, ' || c_version );
  begin
    t_offs_dir_header := dbms_lob.getlength( p_zipped_blob );
    t_offs := 1;
    while dbms_lob.substr( p_zipped_blob, utl_raw.length( c_LOCAL_FILE_HEADER ), t_offs ) = c_LOCAL_FILE_HEADER
    loop
      t_cnt := t_cnt + 1;
      dbms_lob.append( p_zipped_blob
                     , utl_raw.concat( hextoraw( '504B0102' )      -- Central directory file header signature
                                     , hextoraw( '1400' )          -- version 2.0
                                     , dbms_lob.substr( p_zipped_blob, 26, t_offs + 4 )
                                     , hextoraw( '0000' )          -- File comment length
                                     , hextoraw( '0000' )          -- Disk number where file starts
                                     , hextoraw( '0000' )          -- Internal file attributes =>
                                                                   --     0000 binary file
                                                                   --     0100 (ascii)text file
                                     , case
                                         when dbms_lob.substr( p_zipped_blob
                                                             , 1
                                                             , t_offs + 30 + blob2num( p_zipped_blob, 2, t_offs + 26 ) - 1
                                                             ) in ( hextoraw( '2F' ) -- /
                                                                  , hextoraw( '5C' ) -- \
                                                                  )
                                         then hextoraw( '10000000' ) -- a directory/folder
                                         else hextoraw( '2000B681' ) -- a file
                                       end                         -- External file attributes
                                     , little_endian( t_offs - 1 ) -- Relative offset of local file header
                                     , dbms_lob.substr( p_zipped_blob
                                                      , blob2num( p_zipped_blob, 2, t_offs + 26 )
                                                      , t_offs + 30
                                                      )            -- File name
                                     )
                     );
      t_offs := t_offs + 30 + blob2num( p_zipped_blob, 4, t_offs + 18 )  -- compressed size
                            + blob2num( p_zipped_blob, 2, t_offs + 26 )  -- File name length
                            + blob2num( p_zipped_blob, 2, t_offs + 28 ); -- Extra field length
    end loop;
    t_offs_end_header := dbms_lob.getlength( p_zipped_blob );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_END_OF_CENTRAL_DIRECTORY                                -- End of central directory signature
                                   , hextoraw( '0000' )                                        -- Number of this disk
                                   , hextoraw( '0000' )                                        -- Disk where central directory starts
                                   , little_endian( t_cnt, 2 )                                 -- Number of central directory records on this disk
                                   , little_endian( t_cnt, 2 )                                 -- Total number of central directory records
                                   , little_endian( t_offs_end_header - t_offs_dir_header )    -- Size of central directory
                                   , little_endian( t_offs_dir_header )                        -- Offset of start of central directory, relative to start of archive
                                   , little_endian( nvl( utl_raw.length( t_comment ), 0 ), 2 ) -- ZIP file comment length
                                   , t_comment
                                   )
                   );
  end;
--
  function alfan_col( p_col pls_integer )
  return varchar2
  is
  begin
    return case
             when p_col > 702 then chr( 64 + trunc( ( p_col - 27 ) / 676 ) ) || chr( 65 + mod( trunc( ( p_col - 1 ) / 26 ) - 1, 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )
             when p_col > 26  then chr( 64 + trunc( ( p_col - 1 ) / 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )
             else chr( 64 + p_col )
           end;
  end;
  --
  function col_alfan( p_col varchar2 )
  return pls_integer
  is
    l_col varchar2(1000) := rtrim( p_col, '0123456789' );
  begin
    return ascii( substr( l_col, -1 ) ) - 64
         + nvl( ( ascii( substr( l_col, -2, 1 ) ) - 64 ) * 26, 0 )
         + nvl( ( ascii( substr( l_col, -3, 1 ) ) - 64 ) * 676, 0 );
  end;
  --
  procedure clear_workbook
  is
    s pls_integer;
    t_row_ind pls_integer;
  begin
    s := workbook.sheets.first;
    while s is not null
    loop
      t_row_ind := workbook.sheets( s ).rows.first;
      while t_row_ind is not null
      loop
        workbook.sheets( s ).rows( t_row_ind ).delete;
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      end loop;
      workbook.sheets( s ).rows.delete;
      workbook.sheets( s ).widths.delete;
      workbook.sheets( s ).autofilters.delete;
      workbook.sheets( s ).hyperlinks.delete;
      workbook.sheets( s ).col_fmts.delete;
      workbook.sheets( s ).row_fmts.delete;
      workbook.sheets( s ).comments.delete;
      workbook.sheets( s ).mergecells.delete;
      workbook.sheets( s ).validations.delete;
      workbook.sheets( s ).drawings.delete;
      s := workbook.sheets.next( s );
    end loop;
    workbook.strings.delete;
    workbook.str_ind.delete;
    workbook.fonts.delete;
    workbook.fills.delete;
    workbook.borders.delete;
    workbook.numFmts.delete;
    workbook.cellXfs.delete;
    workbook.defined_names.delete;
    workbook.tables.delete;
    for i in 1 .. workbook.images.count
    loop
      dbms_lob.freetemporary( workbook.images(i).img );
    end loop;
    workbook.images.delete;
    workbook := null;
  end;
--
  procedure set_tabcolor
    ( p_tabcolor varchar2 -- this is a hex ALPHA Red Green Blue value
    , p_sheet pls_integer := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( t_sheet ).tabcolor := substr( p_tabcolor, 1, 8 );
  end;
--
  procedure new_sheet
    ( p_sheetname      varchar2    := null
    , p_tabcolor       varchar2    := null -- this is a hex ALPHA Red Green Blue value
    , p_show_gridlines boolean     := null
    , p_grid_color_idx pls_integer := null -- index in default color palette 0 - 55
    , p_show_headers   boolean     := null
    )
  is
    t_nr pls_integer := workbook.sheets.count + 1;
    t_ind pls_integer;
  begin
    workbook.sheets( t_nr ).name := nvl( dbms_xmlgen.convert( translate( p_sheetname, 'a/\[]*:?', 'a' ) ), 'Sheet' || t_nr );
    workbook.sheets( t_nr ).show_gridlines := p_show_gridlines;
    workbook.sheets( t_nr ).grid_color_idx := p_grid_color_idx;
    workbook.sheets( t_nr ).show_headers   := p_show_headers;
    if workbook.strings.count = 0
    then
     workbook.str_cnt := 0;
    end if;
    if workbook.fonts.count = 0
    then
      t_ind := get_font( 'Calibri' );
    end if;
    if workbook.fills.count = 0
    then
      t_ind := get_fill( 'none' );
      t_ind := get_fill( 'gray125' );
    end if;
    if workbook.borders.count = 0
    then
      t_ind := get_border( '', '', '', '' );
    end if;
    set_tabcolor( p_tabcolor, t_nr );
  end;
--
  procedure set_col_width
    ( p_sheet pls_integer
    , p_col pls_integer
    , p_format varchar2
    )
  is
    t_width number;
    t_nr_chr pls_integer;
  begin
    if p_format is null
    then
      return;
    end if;
    if instr( p_format, ';' ) > 0
    then
      t_nr_chr := length( translate( substr( p_format, 1, instr( p_format, ';' ) - 1 ), 'a\"', 'a' ) );
    else
      t_nr_chr := length( translate( p_format, 'a\"', 'a' ) );
    end if;
    t_width := trunc( ( t_nr_chr * 7 + 5 ) / 7 * 256 ) / 256; -- assume default 11 point Calibri
    if workbook.sheets( p_sheet ).widths.exists( p_col )
    then
      workbook.sheets( p_sheet ).widths( p_col ) :=
        greatest( workbook.sheets( p_sheet ).widths( p_col )
                , t_width
                );
    else
      workbook.sheets( p_sheet ).widths( p_col ) := greatest( t_width, 8.43 );
    end if;
  end;
--
  function OraFmt2Excel( p_format varchar2 := null )
  return varchar2
  is
    t_format varchar2(1000) := substr( p_format, 1, 1000 );
  begin
    t_format := replace( replace( t_format, 'hh24', 'hh' ), 'hh12', 'hh' );
    t_format := replace( t_format, 'mi', 'mm' );
    t_format := replace( replace( replace( t_format, 'AM', '~~' ), 'PM', '~~' ), '~~', 'AM/PM' );
    t_format := replace( replace( replace( t_format, 'am', '~~' ), 'pm', '~~' ), '~~', 'AM/PM' );
    t_format := replace( replace( t_format, 'day', 'DAY' ), 'DAY', 'dddd' );
    t_format := replace( replace( t_format, 'dy', 'DY' ), 'DAY', 'ddd' );
    t_format := replace( replace( t_format, 'RR', 'RR' ), 'RR', 'YY' );
    t_format := replace( replace( t_format, 'month', 'MONTH' ), 'MONTH', 'mmmm' );
    t_format := replace( replace( t_format, 'mon', 'MON' ), 'MON', 'mmm' );
    t_format := replace( t_format, '9', '#' );
    t_format := replace( t_format, 'D', '.' );
    t_format := replace( t_format, 'G', ',' );
    return t_format;
  end;
--
  function get_numFmt( p_format varchar2 := null )
  return pls_integer
  is
    t_cnt pls_integer;
    t_numFmtId pls_integer;
  begin
    if p_format is null
    then
      return 0;
    end if;
    t_cnt := workbook.numFmts.count;
    for i in 1 .. t_cnt
    loop
      if workbook.numFmts( i ).formatCode = p_format
      then
        t_numFmtId := workbook.numFmts( i ).numFmtId;
        exit;
      end if;
    end loop;
    if t_numFmtId is null
    then
      t_numFmtId := case when t_cnt = 0 then 164 else workbook.numFmts( t_cnt ).numFmtId + 1 end;
      t_cnt := t_cnt + 1;
      workbook.numFmts( t_cnt ).numFmtId := t_numFmtId;
      workbook.numFmts( t_cnt ).formatCode := p_format;
      workbook.numFmtIndexes( t_numFmtId ) := t_cnt;
    end if;
    return t_numFmtId;
  end;
--
  function get_font
    ( p_name varchar2
    , p_family pls_integer := 2
    , p_fontsize number := 11
    , p_theme pls_integer := 1
    , p_underline boolean := false
    , p_italic boolean := false
    , p_bold boolean := false
    , p_rgb varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
  return pls_integer
  is
    t_ind pls_integer;
  begin
    if workbook.fonts.count > 0
    then
      for f in 0 .. workbook.fonts.count - 1
      loop
        if (   workbook.fonts( f ).name = p_name
           and workbook.fonts( f ).family = p_family
           and workbook.fonts( f ).fontsize = p_fontsize
           and workbook.fonts( f ).theme = p_theme
           and workbook.fonts( f ).underline = p_underline
           and workbook.fonts( f ).italic = p_italic
           and workbook.fonts( f ).bold = p_bold
           and ( workbook.fonts( f ).rgb = p_rgb
               or ( workbook.fonts( f ).rgb is null and p_rgb is null )
               )
           )
        then
          return f;
        end if;
      end loop;
    end if;
    t_ind := workbook.fonts.count;
    workbook.fonts( t_ind ).name := p_name;
    workbook.fonts( t_ind ).family := p_family;
    workbook.fonts( t_ind ).fontsize := p_fontsize;
    workbook.fonts( t_ind ).theme := p_theme;
    workbook.fonts( t_ind ).underline := p_underline;
    workbook.fonts( t_ind ).italic := p_italic;
    workbook.fonts( t_ind ).bold := p_bold;
    workbook.fonts( t_ind ).rgb := p_rgb;
    return t_ind;
  end;
--
  function get_fill
    ( p_patternType varchar2
    , p_fgRGB varchar2 := null
    )
  return pls_integer
  is
    t_ind pls_integer;
  begin
    if workbook.fills.count > 0
    then
      for f in 0 .. workbook.fills.count - 1
      loop
        if (   workbook.fills( f ).patternType = p_patternType
           and nvl( workbook.fills( f ).fgRGB, 'x' ) = nvl( upper( p_fgRGB ), 'x' )
           )
        then
          return f;
        end if;
      end loop;
    end if;
    t_ind := workbook.fills.count;
    workbook.fills( t_ind ).patternType := p_patternType;
    workbook.fills( t_ind ).fgRGB := upper( p_fgRGB );
    return t_ind;
  end;
--
  function get_border
    ( p_top    varchar2 := 'thin'
    , p_bottom varchar2 := 'thin'
    , p_left   varchar2 := 'thin'
    , p_right  varchar2 := 'thin'
    , p_rgb    varchar2 := null
    )
  return pls_integer
  is
    t_ind pls_integer;
  begin
    if workbook.borders.count > 0
    then
      for b in 0 .. workbook.borders.count - 1
      loop
        if (   nvl( workbook.borders( b ).top, 'x' )    = nvl( p_top, 'x' )
           and nvl( workbook.borders( b ).bottom, 'x' ) = nvl( p_bottom, 'x' )
           and nvl( workbook.borders( b ).left, 'x' )   = nvl( p_left, 'x' )
           and nvl( workbook.borders( b ).right, 'x' )  = nvl( p_right, 'x' )
           and nvl( workbook.borders( b ).rgb, 'x' )    = nvl( p_rgb, 'x' )
           )
        then
          return b;
        end if;
      end loop;
    end if;
    t_ind := workbook.borders.count;
    workbook.borders( t_ind ).top    := p_top;
    workbook.borders( t_ind ).bottom := p_bottom;
    workbook.borders( t_ind ).left   := p_left;
    workbook.borders( t_ind ).right  := p_right;
    workbook.borders( t_ind ).rgb    := p_rgb;
    return t_ind;
  end;
--
  function get_alignment
    ( p_vertical   varchar2 := null
    , p_horizontal varchar2 := null
    , p_wrapText   boolean  := null
    )
  return tp_alignment
  is
    t_rv tp_alignment;
  begin
    t_rv.vertical   := p_vertical;
    t_rv.horizontal := p_horizontal;
    t_rv.wrapText   := p_wrapText;
    return t_rv;
  end;
  --
  function get_xfid
    ( p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    )
  return pls_integer
  is
    l_XF   tp_XF_fmt;
    l_cnt  pls_integer;
  begin
    l_XF.numFmtId  := coalesce( p_numFmtId, 0 );
    l_XF.fontId    := coalesce( p_fontId  , 0 );
    l_XF.fillId    := coalesce( p_fillId  , 0 );
    l_XF.borderId  := coalesce( p_borderId, 0 );
    l_XF.alignment := p_alignment;
    if (   l_XF.numFmtId + l_XF.fontId + l_XF.fillId + l_XF.borderId = 0
       and l_XF.alignment.vertical   is null
       and l_XF.alignment.horizontal is null
       and not nvl( l_XF.alignment.wrapText, false )
       )
    then
      return null;
    end if;
    l_cnt := workbook.cellXfs.count;
    for i in 1 .. l_cnt
    loop
      if (   workbook.cellXfs( i ).numFmtId = l_XF.numFmtId
         and workbook.cellXfs( i ).fontId   = l_XF.fontId
         and workbook.cellXfs( i ).fillId   = l_XF.fillId
         and workbook.cellXfs( i ).borderId = l_XF.borderId
         and nvl( workbook.cellXfs( i ).alignment.vertical, 'x' )   = nvl( l_XF.alignment.vertical, 'x' )
         and nvl( workbook.cellXfs( i ).alignment.horizontal, 'x' ) = nvl( l_XF.alignment.horizontal, 'x' )
         and nvl( workbook.cellXfs( i ).alignment.wrapText, false ) = nvl( l_XF.alignment.wrapText, false )
         )
      then
        return i;
      end if;
    end loop;
    l_cnt := l_cnt + 1;
    workbook.cellXfs( l_cnt ) := l_XF;
    return l_cnt;
  end get_XfId;
  --
  function get_XfId
    ( p_sheet     pls_integer
    , p_col       pls_integer
    , p_row       pls_integer
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    )
  return varchar2
  is
    l_XfId   pls_integer;
    t_XF     tp_XF_fmt;
    t_col_XF tp_XF_fmt;
    t_row_XF tp_XF_fmt;
  begin
    if not g_useXf
    then
      return '';
    end if;
    if workbook.sheets( p_sheet ).col_fmts.exists( p_col )
    then
      t_col_XF := workbook.sheets( p_sheet ).col_fmts( p_col );
    end if;
    if workbook.sheets( p_sheet ).row_fmts.exists( p_row )
    then
      t_row_XF := workbook.sheets( p_sheet ).row_fmts( p_row );
    end if;
    t_XF.numFmtId  := coalesce( p_numFmtId, t_col_XF.numFmtId, t_row_XF.numFmtId, 0 );
    t_XF.fontId    := coalesce( p_fontId  , t_col_XF.fontId  , t_row_XF.fontId  , 0 );
    t_XF.fillId    := coalesce( p_fillId  , t_col_XF.fillId  , t_row_XF.fillId  , 0 );
    t_XF.borderId  := coalesce( p_borderId, t_col_XF.borderId, t_row_XF.borderId, 0 );
    t_XF.alignment := get_alignment
                        ( coalesce( p_alignment.vertical, t_col_XF.alignment.vertical, t_row_XF.alignment.vertical )
                        , coalesce( p_alignment.horizontal, t_col_XF.alignment.horizontal, t_row_XF.alignment.horizontal )
                        , coalesce( p_alignment.wrapText, t_col_XF.alignment.wrapText, t_row_XF.alignment.wrapText )
                        );
    l_xfid := get_xfid( t_XF.numFmtId, t_XF.fontId, t_XF.fillId, t_XF.borderId, t_XF.alignment );
    if l_xfid is null
    then
      return '';
    end if;
    if t_XF.numFmtId > 0
    then
      set_col_width( p_sheet, p_col, workbook.numFmts( workbook.numFmtIndexes( t_XF.numFmtId ) ).formatCode );
    end if;
    return 's="' || l_xfid || '"';
  end get_XfId;
  --
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value number
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).value := p_value;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style :=
      case when p_xfid is null
        then get_XfId( t_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment )
        else 's="' || p_xfid || '"'
      end;
  end;
--
  function add_string( p_string varchar2 )
  return pls_integer
  is
    t_cnt pls_integer;
  begin
    if workbook.strings.exists( nvl( p_string, '' ) )
    then
      t_cnt := workbook.strings( nvl( p_string, '' ) );
    else
      t_cnt := workbook.strings.count;
      workbook.str_ind( t_cnt ) := p_string;
      workbook.strings( nvl( p_string, '' ) ) := t_cnt;
    end if;
    workbook.str_cnt := workbook.str_cnt + 1;
    return t_cnt;
  end;
--
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value varchar2
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
    t_alignment tp_alignment := p_alignment;
  begin
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).value := add_string( p_value );
    if t_alignment.wrapText is null and instr( p_value, chr(13) ) > 0
    then
      t_alignment.wrapText := true;
    end if;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style := 't="s" ' ||
      case when p_xfid is null
        then get_XfId( t_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, t_alignment )
        else 's="' || p_xfid || '"'
      end;
  end;
--
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value date
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    )
  is
    t_xfId     varchar2(100);
    t_numFmtId pls_integer := p_numFmtId;
    t_sheet    pls_integer := nvl( p_sheet, workbook.sheets.count );
    t_tmp      number;
  begin
    t_tmp := p_value - date '1900-03-01';
    t_tmp := t_tmp + case when t_tmp < 0 then 60 else 61 end;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).value := t_tmp;
    if p_xfId is null
    then
      if t_numFmtId is null
         and not (   workbook.sheets( t_sheet ).col_fmts.exists( p_col )
                 and workbook.sheets( t_sheet ).col_fmts( p_col ).numFmtId is not null
                 )
         and not (   workbook.sheets( t_sheet ).row_fmts.exists( p_row )
                 and workbook.sheets( t_sheet ).row_fmts( p_row ).numFmtId is not null
                 )
      then
        t_numFmtId := get_numFmt( 'dd/mm/yyyy' );
      end if;
      t_xfId := get_XfId( t_sheet, p_col, p_row, t_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment );
    else
      t_xfId := 's="' || p_xfid || '"';
    end if;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style := t_xfId;
  end;
--
  procedure query_date_cell
    ( p_col pls_integer
    , p_row pls_integer
    , p_value date
    , p_sheet pls_integer := null
    , p_XfId varchar2
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    cell( p_col, p_row, p_value, 0, p_sheet => t_sheet );
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style := p_XfId;
  end;
--
  procedure hyperlink
    ( p_col pls_integer
    , p_row pls_integer
    , p_url varchar2 := null
    , p_value varchar2 := null
    , p_sheet pls_integer := null
    , p_location varchar2 := null
    , p_tooltip varchar2 := null
    )
  is
    t_ind pls_integer;
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    if p_url is not null or p_location is not null
    then
      workbook.sheets( t_sheet ).rows( p_row )( p_col ).value := add_string( coalesce( p_value, p_url, p_location ) );
      workbook.sheets( t_sheet ).rows( p_row )( p_col ).style := 't="s" ' || get_XfId( t_sheet, p_col, p_row, '', get_font( 'Calibri', p_theme => 10, p_underline => true ) );
      t_ind := workbook.sheets( t_sheet ).hyperlinks.count + 1;
      workbook.sheets( t_sheet ).hyperlinks( t_ind ).cell := alfan_col( p_col ) || p_row;
      workbook.sheets( t_sheet ).hyperlinks( t_ind ).url := p_url;
      workbook.sheets( t_sheet ).hyperlinks( t_ind ).location := p_location;
      workbook.sheets( t_sheet ).hyperlinks( t_ind ).tooltip := p_tooltip;
    end if;
  end;
--
  procedure comment
    ( p_col pls_integer
    , p_row pls_integer
    , p_text varchar2
    , p_author varchar2 := null
    , p_width pls_integer := 150
    , p_height pls_integer := 100
    , p_sheet pls_integer := null
    )
  is
    t_ind pls_integer;
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    t_ind := workbook.sheets( t_sheet ).comments.count + 1;
    workbook.sheets( t_sheet ).comments( t_ind ).row := p_row;
    workbook.sheets( t_sheet ).comments( t_ind ).column := p_col;
    workbook.sheets( t_sheet ).comments( t_ind ).text := dbms_xmlgen.convert( p_text );
    workbook.sheets( t_sheet ).comments( t_ind ).author := dbms_xmlgen.convert( p_author );
    workbook.sheets( t_sheet ).comments( t_ind ).width := p_width;
    workbook.sheets( t_sheet ).comments( t_ind ).height := p_height;
  end;
--
  procedure mergecells
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_sheet pls_integer := null
    )
  is
    t_ind pls_integer;
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    t_ind := workbook.sheets( t_sheet ).mergecells.count + 1;
    workbook.sheets( t_sheet ).mergecells( t_ind ) := alfan_col( p_tl_col ) || p_tl_row || ':' || alfan_col( p_br_col ) || p_br_row;
  end;
--
  procedure add_validation
    ( p_type varchar2
    , p_sqref varchar2
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_formula1 varchar2 := null
    , p_formula2 varchar2 := null
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    )
  is
    t_ind pls_integer;
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    t_ind := workbook.sheets( t_sheet ).validations.count + 1;
    workbook.sheets( t_sheet ).validations( t_ind ).type := p_type;
    workbook.sheets( t_sheet ).validations( t_ind ).errorstyle := p_style;
    workbook.sheets( t_sheet ).validations( t_ind ).sqref := p_sqref;
    workbook.sheets( t_sheet ).validations( t_ind ).formula1 := p_formula1;
    workbook.sheets( t_sheet ).validations( t_ind ).error_title := p_error_title;
    workbook.sheets( t_sheet ).validations( t_ind ).error_txt := p_error_txt;
    workbook.sheets( t_sheet ).validations( t_ind ).title := p_title;
    workbook.sheets( t_sheet ).validations( t_ind ).prompt := p_prompt;
    workbook.sheets( t_sheet ).validations( t_ind ).showerrormessage := p_show_error;
  end;
--
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    )
  is
  begin
    add_validation( 'list'
                  , alfan_col( p_sqref_col ) || p_sqref_row
                  , p_style => lower( p_style )
                  , p_formula1 => '$' || alfan_col( p_tl_col ) || '$' ||  p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row
                  , p_title => p_title
                  , p_prompt => p_prompt
                  , p_show_error => p_show_error
                  , p_error_title => p_error_title
                  , p_error_txt => p_error_txt
                  , p_sheet => p_sheet
                  );
  end;
--
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_defined_name varchar2
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    )
  is
  begin
    add_validation( 'list'
                  , alfan_col( p_sqref_col ) || p_sqref_row
                  , p_style => lower( p_style )
                  , p_formula1 => p_defined_name
                  , p_title => p_title
                  , p_prompt => p_prompt
                  , p_show_error => p_show_error
                  , p_error_title => p_error_title
                  , p_error_txt => p_error_txt
                  , p_sheet => p_sheet
                  );
  end;
--
  procedure defined_name
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_name varchar2
    , p_sheet pls_integer := null
    , p_localsheet pls_integer := null
    )
  is
    t_ind pls_integer;
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    t_ind := workbook.defined_names.count + 1;
    workbook.defined_names( t_ind ).name := p_name;
    workbook.defined_names( t_ind ).ref := 'Sheet' || t_sheet || '!$' || alfan_col( p_tl_col ) || '$' ||  p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row;
    workbook.defined_names( t_ind ).sheet := p_localsheet;
  end;
--
  procedure set_column_width
    ( p_col pls_integer
    , p_width number
    , p_sheet pls_integer := null
    )
  is
    t_width number;
  begin
    t_width := trunc( round( p_width * 7 ) * 256 / 7 ) / 256;
    workbook.sheets( nvl( p_sheet, workbook.sheets.count ) ).widths( p_col ) := t_width;
  end;
--
  procedure set_column
    ( p_col pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( t_sheet ).col_fmts( p_col ).numFmtId := p_numFmtId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).fontId := p_fontId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).fillId := p_fillId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).borderId := p_borderId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).alignment := p_alignment;
  end;
--
  procedure set_row
    ( p_row pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    , p_height number := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
    t_cells tp_cells;
  begin
    workbook.sheets( t_sheet ).row_fmts( p_row ).numFmtId := p_numFmtId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).fontId := p_fontId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).fillId := p_fillId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).borderId := p_borderId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).alignment := p_alignment;
    workbook.sheets( t_sheet ).row_fmts( p_row ).height := trunc( p_height * 4 / 3 ) * 3 / 4;
    if not workbook.sheets( t_sheet ).rows.exists( p_row )
    then
      workbook.sheets( t_sheet ).rows( p_row ) := t_cells;
    end if;
  end;
--
  procedure freeze_rows
    ( p_nr_rows pls_integer := 1
    , p_sheet pls_integer := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( t_sheet ).freeze_cols := null;
    workbook.sheets( t_sheet ).freeze_rows := p_nr_rows;
  end;
--
  procedure freeze_cols
    ( p_nr_cols pls_integer := 1
    , p_sheet pls_integer := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( t_sheet ).freeze_rows := null;
    workbook.sheets( t_sheet ).freeze_cols := p_nr_cols;
  end;
--
  procedure freeze_pane
    ( p_col pls_integer
    , p_row pls_integer
    , p_sheet pls_integer := null
    )
  is
    t_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( t_sheet ).freeze_rows := p_row;
    workbook.sheets( t_sheet ).freeze_cols := p_col;
  end;
--
  procedure set_autofilter
    ( p_column_start pls_integer := null
    , p_column_end pls_integer := null
    , p_row_start pls_integer := null
    , p_row_end pls_integer := null
    , p_sheet pls_integer := null
    )
  is
    t_ind pls_integer;
    t_sheet pls_integer := coalesce( p_sheet, workbook.sheets.count );
  begin
    t_ind := 1;
    workbook.sheets( t_sheet ).autofilters( t_ind ).column_start := p_column_start;
    workbook.sheets( t_sheet ).autofilters( t_ind ).column_end := p_column_end;
    workbook.sheets( t_sheet ).autofilters( t_ind ).row_start := p_row_start;
    workbook.sheets( t_sheet ).autofilters( t_ind ).row_end := p_row_end;
    defined_name
      ( p_column_start
      , p_row_start
      , p_column_end
      , p_row_end
      , '_xlnm._FilterDatabase'
      , t_sheet
      , t_sheet - 1
      );
  end;
--
  procedure set_table
    ( p_column_start pls_integer
    , p_column_end   pls_integer
    , p_row_start    pls_integer
    , p_row_end      pls_integer
    , p_style        varchar2
    , p_name         varchar2    := null
    , p_sheet        pls_integer := null
    )
  is
    t_table tp_table;
    t_cnt pls_integer := workbook.tables.count + 1;
  begin
    t_table.sheet := coalesce( p_sheet, workbook.sheets.count );
    t_table.column_start := p_column_start;
    t_table.column_end   := p_column_end;
    t_table.row_start    := p_row_start;
    t_table.row_end      := p_row_end;
    t_table.name         := coalesce( p_name, 'Table' || t_cnt );
    t_table.style        := p_style;
    workbook.tables( t_cnt ) := t_table;
  end;
--
/*
  procedure add1xml
    ( p_excel in out nocopy blob
    , p_filename varchar2
    , p_xml clob
    )
  is
    t_tmp blob;
    c_step constant number := 24396;
  begin
    dbms_lob.createtemporary( t_tmp, true );
    for i in 0 .. trunc( length( p_xml ) / c_step )
    loop
      dbms_lob.append( t_tmp, utl_i18n.string_to_raw( substr( p_xml, i * c_step + 1, c_step ), 'AL32UTF8' ) );
    end loop;
    add1file( p_excel, p_filename, t_tmp );
    dbms_lob.freetemporary( t_tmp );
  end;
*/
--
  procedure add1xml
    ( p_excel in out nocopy blob
    , p_filename varchar2
    , p_xml clob
    )
  is
    t_tmp blob;
    dest_offset integer := 1;
    src_offset integer := 1;
    lang_context integer;
    warning integer;
  begin
    lang_context := dbms_lob.DEFAULT_LANG_CTX;
    dbms_lob.createtemporary( t_tmp, true );
    dbms_lob.converttoblob
      ( t_tmp
      , p_xml
      , dbms_lob.lobmaxsize
      , dest_offset
      , src_offset
      ,  nls_charset_id( 'AL32UTF8'  )
      , lang_context
      , warning
      );
    add1file( p_excel, p_filename, t_tmp );
    dbms_lob.freetemporary( t_tmp );
  end;
--
  function finish_drawing( p_drawing tp_drawing, p_idx pls_integer, p_sheet pls_integer )
  return varchar2
  is
    t_rv varchar2(32767);
    t_col pls_integer;
    t_row pls_integer;
    t_width number;
    t_height number;
    t_col_offs number;
    t_row_offs number;
    t_col_width number;
    t_row_height number;
    t_widths tp_widths;
    t_heights tp_row_fmts;
  begin
    t_width  := workbook.images( p_drawing.img_id ).width;
    t_height := workbook.images( p_drawing.img_id ).height;
    if p_drawing.scale is not null
    then
      t_width  := p_drawing.scale * t_width;
      t_height := p_drawing.scale * t_height;
    end if;
    if workbook.sheets( p_sheet ).widths.count = 0
    then
-- assume default column widths!
-- 64 px = 1 col = 609600
      t_col := trunc( t_width / 64 );
      t_col_offs := ( t_width - t_col * 64 ) * 9525;
      t_col := p_drawing.col - 1 + t_col;
    else
      t_widths := workbook.sheets( p_sheet ).widths;
      t_col := p_drawing.col;
      loop
        if t_widths.exists( t_col )
        then
          t_col_width := round( 7 * t_widths( t_col ) );
        else
          t_col_width := 64;
        end if;
        exit when t_width < t_col_width;
        t_col := t_col + 1;
        t_width := t_width - t_col_width;
      end loop;
      t_col := t_col - 1;
      t_col_offs := t_width * 9525;
    end if;
--
    if workbook.sheets( p_sheet ).row_fmts.count = 0
    then
-- assume default row heigths!
-- 20 px = 1 row = 190500
      t_row := trunc( t_height / 20 );
      t_row_offs := ( t_height - t_row * 20 ) * 9525;
      t_row := p_drawing.row - 1 + t_row;
    else
      t_heights := workbook.sheets( p_sheet ).row_fmts;
      t_row := p_drawing.row;
      loop
        if t_heights.exists( t_row ) and t_heights( t_row ).height is not null
        then
          t_row_height := t_heights( t_row ).height;
          t_row_height := round( 4 * t_row_height / 3 );
        else
          t_row_height := 20;
        end if;
        exit when t_height < t_row_height;
        t_row := t_row + 1;
        t_height := t_height - t_row_height;
      end loop;
      t_row_offs := t_height * 9525;
      t_row := t_row - 1;
    end if;
    t_rv := '<xdr:twoCellAnchor editAs="oneCell">
<xdr:from>
<xdr:col>' || ( p_drawing.col - 1 ) || '</xdr:col>
<xdr:colOff>0</xdr:colOff>
<xdr:row>' || ( p_drawing.row - 1 ) || '</xdr:row>
<xdr:rowOff>0</xdr:rowOff>
</xdr:from>
<xdr:to>
<xdr:col>' || t_col || '</xdr:col>
<xdr:colOff>' || t_col_offs || '</xdr:colOff>
<xdr:row>' || t_row || '</xdr:row>
<xdr:rowOff>' || t_row_offs || '</xdr:rowOff>
</xdr:to>
<xdr:pic>
<xdr:nvPicPr>
<xdr:cNvPr id="3" name="' || coalesce( p_drawing.name, 'Picture ' || p_idx ) || '"';
    if p_drawing.title is not null
    then
      t_rv := t_rv || ' title="' || p_drawing.title || '"';
    end if;
    if p_drawing.description is not null
    then
      t_rv := t_rv || ' descr="' || p_drawing.description || '"';
    end if;
    t_rv := t_rv || '/>
<xdr:cNvPicPr>
<a:picLocks noChangeAspect="1"/>
</xdr:cNvPicPr>
</xdr:nvPicPr>
<xdr:blipFill>
<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' || p_drawing.img_id || '">
<a:extLst>
<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
</a:ext>
</a:extLst>
</a:blip>
<a:stretch>
<a:fillRect/>
</a:stretch>
</xdr:blipFill>
<xdr:spPr>
<a:prstGeom prst="rect">
</a:prstGeom>
</xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:twoCellAnchor>
';
    return t_rv;
  end;
  --
  function add_rgb( p_rgb varchar2, p_tag varchar2 := 'color' )
  return varchar2
  is
  begin
    return case when p_rgb is not null then '<' || p_tag || ' rgb="' || p_rgb || '"/>' end;
  end;
  --
$IF as_xlsx.use_dbms_crypto
$THEN
  function excel_encrypt( p_xlsx blob, p_password varchar2 )
  return blob
  is
    t_EncryptionInfo raw(32767);
    t_EncryptedPackage blob;
    --
    c_Free_SecID         constant pls_integer := -1; -- Free sector, may exist in the file, but is not part of any stream
    c_End_Of_Chain_SecID constant pls_integer := -2; -- Trailing SecID in a SecID chain
    c_SAT_SecID          constant pls_integer := -3; -- Sector is used by the sector allocation table
    c_MSAT_SecID         constant pls_integer := -4; -- Sector is used by the master sector allocation table
    --
    c_CLR_Red      constant raw(1) := hextoraw( '00' ); -- Red
    c_CLR_Black    constant raw(1) := hextoraw( '01' ); -- Black
    --
    c_DIR_Empty    constant raw(1) := hextoraw( '00' ); -- Empty
    c_DIR_Storage  constant raw(1) := hextoraw( '01' ); -- User storage
    c_DIR_Stream   constant raw(1) := hextoraw( '02' ); -- User stream
    c_DIR_Lock     constant raw(1) := hextoraw( '03' ); -- LockBytes
    c_DIR_Property constant raw(1) := hextoraw( '04' ); -- Property
    c_DIR_Root     constant raw(1) := hextoraw( '05' ); -- Root storage
    --
    c_Primary raw(200) := hextoraw( '58000000010000004C0000007B00460046003900410033004600300033002D0035003600450046002D0034003600310033002D0042004400440035002D003500410034003100430031004400300037003200340036007D004E0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E0045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D00000001000000010000000100000000000000000000000000000004000000' );
    c_StrongEncryptionDataSpace raw(64) := hextoraw( '0800000001000000320000005300740072006F006E00670045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D000000' );
    c_DataSpaceMap raw(112) := hextoraw( '08000000010000006800000001000000000000002000000045006E0063007200790070007400650064005000610063006B00610067006500320000005300740072006F006E00670045006E006300720079007000740069006F006E004400610074006100530070006100630065000000' );
    c_Version raw(76) := hextoraw( '3C0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E004400610074006100530070006100630065007300010000000100000001000000' );
    --
    type tp_childs is table of pls_integer index by pls_integer;
    type tp_dir_entry is record
      ( rname raw(64)
      , tp_entry raw(1)
      , colour   raw(1) := c_CLR_Red
      , left  pls_integer := -1
      , right pls_integer := -1
      , root  pls_integer := -1
      , childs tp_childs
      , len pls_integer := 0
      , first_sector pls_integer := 0
      );
    type tp_dir is table of tp_dir_entry index by pls_integer;
    t_dir tp_dir;
    t_cf blob;
    t_short_stream blob;
    t_root     pls_integer;
    t_dummy    pls_integer;
    t_storage  pls_integer;
    t_storage2 pls_integer;
    t_sorted boolean;
    t_tmp pls_integer;
    t_header raw(512);
    t_ssz        pls_integer := 512; -- sector size
    t_sssz       pls_integer := 64;  -- short sector size
    t_ss_cutoff  pls_integer := 4096;
    t_sectId     pls_integer;
    t_tmp_sectId t_sectId%type;
    type tp_secids is table of t_sectId%type index by pls_integer;
    t_msat tp_secids;
    t_sat  tp_secids;
    t_ssat tp_secids;
    t_st_dir  pls_integer;
    t_st_ssf  pls_integer;
    t_cnt_ssf pls_integer;
    --
    function is_less( p1 tp_dir_entry, p2 tp_dir_entry )
    return boolean
    is
    begin
      return case sign( utl_raw.length( p1.rname ) - utl_raw.length( p2.rname ) )
               when -1 then true
               when  1 then false
               else upper( utl_i18n.raw_to_char( p1.rname, 'AL16UTF16LE' ) )
                  < upper( utl_i18n.raw_to_char( p2.rname, 'AL16UTF16LE' ) )
             end;
    end;
  --
  function add_dir_entry( p_name varchar2, tp_entry raw, p_parent pls_integer := null, p_stream blob := null, p_prefix raw := null )
  return pls_integer
  is
    t_id pls_integer;
    t_entry tp_dir_entry;
  begin
    t_id := t_dir.count;
    t_entry.tp_entry := tp_entry;
    t_entry.rname := utl_raw.concat( p_prefix, utl_i18n.string_to_raw( p_name, 'AL16UTF16LE' ) );
    if p_parent is not null
    then
      t_dir( p_parent ).childs( t_dir( p_parent ).childs.count ) := t_id;
    end if;
    if tp_entry = c_DIR_Stream
    then
      t_entry.len := dbms_lob.getlength( p_stream );
      if t_entry.len >= t_ss_cutoff
      then
        dbms_lob.append( t_cf, p_stream );
        if mod( t_entry.len, t_ssz ) > 0
        then
          dbms_lob.writeappend( t_cf, t_ssz - mod( t_entry.len, t_ssz ), utl_raw.copies( '00', t_ssz ) );
        end if;
        t_entry.first_sector := t_sat.count;
        for i in t_sat.count .. t_sat.count + trunc( ( t_entry.len - 1 ) / t_ssz ) - 1
        loop
          t_sat( i ) := i + 1;
        end loop;
        t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
      else
        dbms_lob.append( t_short_stream, p_stream );
        if mod( t_entry.len, t_sssz ) > 0
        then
          dbms_lob.writeappend( t_short_stream, t_sssz - mod( t_entry.len, t_sssz ), utl_raw.copies( '00', t_sssz ) );
        end if;
        t_entry.first_sector := t_ssat.count;
        for i in t_ssat.count .. t_ssat.count + trunc( ( t_entry.len - 1 ) / t_sssz ) - 1
        loop
          t_ssat( i ) := i + 1;
        end loop;
        t_ssat( t_ssat.count ) := c_End_Of_Chain_SecID;
      end if;
    end if;
    t_dir( t_id ) := t_entry;
    return t_id;
  end;
--
  procedure doEncryption
    ( p_pw      varchar2
    , p_excel   blob
    , p_package in out blob
    , p_info    in out raw
    )
  is
    c_algo      constant pls_integer   := dbms_crypto.ENCRYPT_AES + dbms_crypto.CHAIN_CBC + dbms_crypto.PAD_ZERO;
    c_keybits   constant pls_integer   := 256 / 8;
    c_hash      constant pls_integer   := dbms_crypto.hash_sh1;
    c_hmac      constant pls_integer   := dbms_crypto.hmac_sh1;
    c_hash_algo constant varchar2(10)  := 'SHA1';
    c_hash_len  constant pls_integer   := utl_raw.length( dbms_crypto.hash( '00', c_hash ) );
    blockSize pls_integer := 16;
    c_spinCount constant pls_integer := 1000;
    c_saltSize  constant pls_integer   := 16;
    c_salt      constant raw(3999)     := dbms_crypto.randombytes( c_saltsize );
    c_data_salt constant raw(3999)     := dbms_crypto.randombytes( c_saltsize );
    c_pw        constant raw(32767)    := utl_i18n.string_to_raw( p_pw, 'AL16UTF16LE' );
    --
    encrVerifierHashInputBlockKey  constant raw(8) :=  hextoraw( 'fea7d2763b4b9e79' );
    encrVerifierHashValueBlockKey  constant raw(8) :=  hextoraw( 'd7aa0f6d3061344e' );
    encryptedKeyValueBlockKey      constant raw(8) :=  hextoraw( '146e0be7abacd0d6' );
    encrIntegritySaltBlockKey      constant raw(8) :=  hextoraw( '5fb2ad010cb9e1f6' );
    encrIntegrityHmacValueBlocKkey constant raw(8) :=  hextoraw( 'a0677f02b22c8433' );
    --
    l_len        integer;
    l_last_block pls_integer;
  decryptedVerifierInput raw(100);
  decryptedVerifierValue raw(100);
  verifierInputKey raw(100);
  verifierValueKey raw(100);
    encryptedKeyValue      varchar2(100);
    encryptedHmacKey       varchar2(100);
    encryptedHmacValue     varchar2(100);
    encryptedVerifierInput varchar2(100);
    encryptedVerifierValue varchar2(100);
    saltRaw           raw(100);
    hashRaw           raw(100);
    ivRaw             raw(100);
    mac               raw(100);
    rkey              raw(100);
    rinp              raw(100);
    decryptedKeyValue raw(100);
    t_block raw(4096);
--
    function GenerateKey( salt raw, password raw, blockKey raw, hashSize number )
    return raw
    is
      hashBuf raw(1000);
    begin
      hashBuf := dbms_crypto.hash( utl_raw.concat( salt, password ), c_hash );
      for i in 0 .. c_spinCount - 1
      loop
        hashBuf := dbms_crypto.hash( utl_raw.concat( little_endian( i ), hashBuf ), c_hash );
      end loop;
      hashBuf := dbms_crypto.hash( utl_raw.concat( hashBuf, blockKey ), c_hash );
      if c_hash_len < hashSize
      then
        hashBuf := utl_raw.concat( hashBuf, utl_raw.copies( hextoraw( '36' ), hashSize ) );
      end if;
      return utl_raw.substr( hashBuf, 1, hashSize );
    end GenerateKey;
  begin
    decryptedKeyValue := dbms_crypto.randombytes( c_keybits );
    rkey := GenerateKey( c_salt, c_pw, encryptedKeyValueBlockKey, c_keybits );
    ivRaw := dbms_crypto.encrypt( decryptedKeyValue, c_algo, rkey, c_salt );
    encryptedKeyValue := utl_raw.cast_to_varchar2( utl_encode.base64_encode( ivRaw ) );
    l_len := dbms_lob.getlength( p_excel );
    p_package := little_endian( l_len, 8 );
    l_last_block := trunc( ( l_len - 1 ) / 4096 );
    for i in 0 .. l_last_block
    loop
      ivRaw := dbms_crypto.hash( utl_raw.concat( c_data_salt, little_endian( i ) ), c_hash );
      if c_hash_len < blockSize
      then
        ivRaw := utl_raw.concat( ivRaw, utl_raw.copies( hextoraw( '36' ), blockSize ) );
      end if;
      ivRaw := utl_raw.substr( ivRaw, 1, blockSize );
      t_block := dbms_lob.substr( p_excel, 4096, 1 + i * 4096 );
      if i = l_last_block and mod( utl_raw.length( t_block ), blockSize ) != 0
      then
        t_block := utl_raw.concat( t_block, utl_raw.copies( 'FF', blockSize - mod( utl_raw.length( t_block ), blockSize ) ) );
      end if;
      dbms_lob.append( p_package, dbms_crypto.encrypt( t_block, c_algo, decryptedKeyValue, ivRaw ) );
    end loop;
--
    saltRaw := dbms_crypto.randombytes( c_hash_len );
    mac := dbms_crypto.mac( p_package, c_hmac, saltRaw );
    ivRaw := dbms_crypto.hash( utl_raw.concat( c_data_salt, encrIntegritySaltBlockKey ), c_hash );
    if utl_raw.length( ivRaw ) < blockSize
    then
      ivRaw := utl_raw.concat( ivRaw, utl_raw.copies( hextoraw( '00' ), blockSize ) );
    end if;
    ivRaw := utl_raw.substr( ivRaw, 1, blockSize );
    saltRaw := dbms_crypto.encrypt( saltRaw, c_algo, decryptedKeyValue, ivRaw );
    encryptedHmacKey := utl_raw.cast_to_varchar2( utl_encode.base64_encode( saltRaw ) );
    ivRaw := dbms_crypto.hash( utl_raw.concat( c_data_salt, encrIntegrityHmacValueBlocKkey ), c_hash );
    if utl_raw.length( ivRaw ) < blockSize
    then
      ivRaw := utl_raw.concat( ivRaw, utl_raw.copies( hextoraw( '00' ), blockSize ) );
    end if;
    ivRaw := utl_raw.substr( ivRaw, 1, blockSize );
    hashRaw := dbms_crypto.encrypt( mac, c_algo, decryptedKeyValue, ivRaw );
    encryptedHmacValue := utl_raw.cast_to_varchar2( utl_encode.base64_encode( hashRaw ) );
--
    rinp := dbms_crypto.randombytes( c_saltSize );
    rkey := GenerateKey( c_salt, c_pw, encrVerifierHashInputBlockKey, c_keybits );
    hashRaw := dbms_crypto.encrypt( rinp, c_algo, rkey, c_salt );
    encryptedVerifierInput := utl_raw.cast_to_varchar2( utl_encode.base64_encode( hashRaw ) );
    rkey := GenerateKey( c_salt, c_pw, encrVerifierHashValueBlockKey, c_keybits );
    rinp := dbms_crypto.hash( rinp, c_hash );
    hashRaw := dbms_crypto.encrypt( rinp, c_algo, rkey, c_salt );
    encryptedVerifierValue := utl_raw.cast_to_varchar2( utl_encode.base64_encode( hashRaw ) );
--
    p_info := utl_raw.concat( hextoraw( '0400040040000000' )
                            , utl_raw.cast_to_raw(
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' || chr(13) || chr(10) ||
'<encryption' ||
' xmlns="http://schemas.microsoft.com/office/2006/encryption"' ||
' xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData' ||
' saltSize="' || to_char( c_saltsize ) || '"' ||
' blockSize="' || to_char( blocksize ) || '"' ||
' keyBits="' || to_char( c_keybits * 8 )|| '"' ||
' hashSize="' || to_char( c_hash_len ) || '"' ||
' cipherAlgorithm="AES"' ||
' cipherChaining="ChainingModeCBC"' ||
' hashAlgorithm="' || c_hash_algo || '"' ||
' saltValue="' || utl_raw.cast_to_varchar2( utl_encode.base64_encode( c_data_salt ) ) || '"/><dataIntegrity' ||
' encryptedHmacKey="' || encryptedHmacKey || '"' ||
' encryptedHmacValue="' || encryptedHmacValue || '"/><keyEncryptors><keyEncryptor' ||
' uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey' ||
' spinCount="' || to_char( c_spincount ) || '"' ||
' saltSize="' || to_char( c_saltsize ) || '"' ||
' blockSize="' || to_char( blocksize ) || '"' ||
' keyBits="' || to_char( c_keybits * 8 ) || '"' ||
' hashSize="' || to_char( c_hash_len ) || '"' ||
' cipherAlgorithm="AES"' ||
' cipherChaining="ChainingModeCBC"' ||
' hashAlgorithm="' || c_hash_algo || '"' ||
' saltValue="' || utl_raw.cast_to_varchar2( utl_encode.base64_encode( c_salt ) ) || '"' ||
' encryptedVerifierHashInput="' || encryptedVerifierInput || '"' ||
' encryptedVerifierHashValue="' || encryptedVerifierValue || '"' ||
' encryptedKeyValue="' || encryptedKeyValue  || '"/></keyEncryptor></keyEncryptors></encryption>' ));
  end doEncryption;
  begin
    doEncryption( p_password, p_xlsx, t_EncryptedPackage, t_EncryptionInfo );
    --
    t_cf := utl_raw.copies( '00', 512 );
    dbms_lob.createtemporary( t_short_stream, true );
    t_root := add_dir_entry( 'Root Entry', c_DIR_Root );
    t_dummy := add_dir_entry( 'EncryptedPackage', c_DIR_Stream, t_root, t_EncryptedPackage );
    t_storage := add_dir_entry( 'DataSpaces', c_DIR_Storage, t_root, p_prefix => '0600' );
    t_dummy := add_dir_entry( 'Version', c_DIR_Stream, t_storage, c_version );
    t_dummy := add_dir_entry( 'DataSpaceMap', c_DIR_Stream, t_storage, c_DataSpaceMap );
    t_storage2 := add_dir_entry( 'DataSpaceInfo', c_DIR_Storage,t_storage );
    t_dummy := add_dir_entry( 'StrongEncryptionDataSpace', c_DIR_Stream, t_storage2, c_StrongEncryptionDataSpace );
    t_storage2 := add_dir_entry( 'TransformInfo', c_DIR_Storage, t_storage );
    t_storage2 := add_dir_entry( 'StrongEncryptionTransform', c_DIR_Storage, t_storage2 );
    t_dummy := add_dir_entry( 'Primary', c_DIR_Stream, t_storage2, c_Primary, p_prefix => '0600' );
    t_dummy := add_dir_entry( 'EncryptionInfo', c_DIR_Stream, t_root, t_EncryptionInfo );
    --
    dbms_lob.freetemporary( t_EncryptedPackage );
    --
    -- write the short sector stream
    dbms_lob.append( t_cf, t_short_stream );
    if mod( dbms_lob.getlength( t_short_stream ), t_ssz ) > 0
    then
      dbms_lob.writeappend( t_cf, t_ssz - mod( dbms_lob.getlength( t_short_stream ), t_ssz ), utl_raw.copies( '00', t_ssz ) );
    end if;
    t_dir( 0 ).len := dbms_lob.getlength( t_short_stream );
    t_dir( 0 ).first_sector := t_sat.count;
    for i in t_sat.count .. t_sat.count + trunc( ( dbms_lob.getlength( t_short_stream ) - 1 ) / t_ssz ) - 1
    loop
      t_sat( i ) := i + 1;
    end loop;
    t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
    --
    -- write the ssat
    for i in 0 .. t_ssat.count - 1
    loop
      dbms_lob.writeappend( t_cf, 4, little_endian( t_ssat( i ) ) );
    end loop;
    if mod( t_ssat.count * 4, t_ssz ) > 0
    then
      dbms_lob.writeappend( t_cf, t_ssz - mod( t_ssat.count * 4, t_ssz ), utl_raw.copies( little_endian( c_Free_SecID ), t_ssz ) );
    end if;
    t_st_ssf := t_sat.count;
    for i in t_sat.count .. t_sat.count + trunc( ( t_ssat.count * 4 - 1 ) / t_ssz ) - 1
    loop
      t_sat( i ) := i + 1;
    end loop;
    t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
    t_cnt_ssf := t_sat.count - t_st_ssf;
    --
    for i in 0 .. t_dir.last
    loop
      if t_dir( i ).childs.count = 1
      then
        t_dir( i ).root := t_dir( i ).childs( 0 );
        t_dir( t_dir( i ).childs( 0 ) ).colour := c_CLR_Black;
      elsif t_dir( i ).childs.count > 1
      then
        t_sorted := false;
        while not t_sorted
        loop
          t_sorted := true;
          for j in 0 .. t_dir( i ).childs.count - 2
          loop
            if is_less( t_dir( t_dir( i ).childs( j + 1 ) ), t_dir( t_dir( i ).childs( j ) ) )
            then
              t_tmp := t_dir( i ).childs( j ) ;
              t_dir( i ).childs( j ) := t_dir( i ).childs( j + 1 );
              t_dir( i ).childs( j + 1) := t_tmp;
              t_sorted := false;
            end if;
          end loop;
        end loop;
        --
        t_tmp := t_dir( i ).childs( 1 );
        t_dir( i ).root := t_tmp;
        t_dir( t_tmp ).left := t_dir( i ).childs( 0 );
        t_dir( t_tmp ).colour := c_CLR_Black;
        if t_dir( i ).childs.count > 2
        then
          t_dir( t_tmp ).right := t_dir( i ).childs( 2 );
          if t_dir( i ).childs.count > 3
          then
            t_dir( t_dir( i ).childs( 2 ) ).right := t_dir( i ).childs( 3 );
            t_dir( t_dir( i ).childs( 0 ) ).colour := c_CLR_Black;
            t_dir( t_dir( i ).childs( 2 ) ).colour := c_CLR_Black;
          end if;
        end if;
      end if;
    end loop;
    --
    -- write the dir tree
    for i in 0 .. t_dir.count - 1
    loop
      dbms_lob.writeappend( t_cf, 128
                          , utl_raw.concat( utl_raw.overlay( '00', t_dir( i ).rname, 64 )
                                          , little_endian( utl_raw.length( t_dir( i ).rname ) + 2, 2 )
                                          , t_dir( i ).tp_entry
                                          , t_dir( i ).colour
                                          , little_endian( t_dir( i ).left )
                                          , little_endian( t_dir( i ).right )
                                          , little_endian( t_dir( i ).root )
                                          , utl_raw.copies( '00', 36 )
                                          , little_endian( t_dir( i ).first_sector )
                                          , little_endian( t_dir( i ).len )
                                          , utl_raw.copies( '00', 4 )
                                          )
                          );
      end loop;
      if mod( t_dir.count * 128, t_ssz ) > 0
      then
        dbms_lob.writeappend( t_cf, t_ssz - mod( t_dir.count * 128, t_ssz ), utl_raw.copies( '00', t_ssz ) );
      end if;
      t_st_dir := t_sat.count;
      for i in t_st_dir .. t_st_dir + trunc( ( t_dir.count * 128 - 1 ) / t_ssz ) - 1
      loop
        t_sat( i ) := i + 1;
      end loop;
      t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
      --
      -- write the sat
      t_tmp := floor( t_sat.count * 4 / t_ssz );
      for i in 0 .. t_tmp
      loop
        t_msat( t_msat.count ) := t_sat.count;
        t_sat( t_sat.count ) := c_SAT_SecID;
      end loop;
      if t_tmp != floor( t_sat.count * 4 / t_ssz )
      then
        t_msat( t_msat.count ) := t_sat.count;
        t_sat( t_sat.count ) := c_SAT_SecID;
      end if;
      for i in 0 .. t_sat.count - 1
      loop
        dbms_lob.writeappend( t_cf, 4, little_endian( t_sat( i ) ) );
      end loop;
      if mod( t_sat.count * 4, t_ssz ) > 0
      then
        dbms_lob.writeappend( t_cf, t_ssz - mod( t_sat.count * 4, t_ssz ), utl_raw.copies( little_endian( c_Free_SecID ), t_ssz ) );
      end if;
      t_header := utl_raw.concat( hextoraw( 'D0CF11E0A1B11AE1' )
                                , utl_raw.copies( '00', 16 )
                                , hextoraw( '3E000300' )
                                , hextoraw( 'FEFF' )
                                , little_endian( round( log( 2, t_ssz ) ), 2 )
                                , little_endian( round( log( 2, t_sssz ) ), 2 )
                                , utl_raw.copies( '00', 10 )
                                , little_endian( t_msat.count )
                                , little_endian( t_st_dir )
                                , utl_raw.copies( '00', 4 )
                                , little_endian( t_ss_cutoff )
                                , little_endian( t_st_ssf )
                                );
    t_header := utl_raw.concat( t_header
                              , little_endian( t_cnt_ssf )
                              , little_endian( c_End_Of_Chain_SecID )
                              , utl_raw.copies( '00', 4 )
                              );
    for i in 0 .. t_msat.count - 1
    loop
      t_header := utl_raw.concat( t_header
                                , little_endian( t_msat( i ) )
                                );
    end loop;
    t_header := utl_raw.concat( t_header
                              , utl_raw.copies( little_endian( c_Free_SecID ), 109 - t_msat.count )
                              );
    dbms_lob.copy( t_cf, t_header, 512, 1, 1 );
    dbms_lob.freetemporary( t_short_stream );
    return t_cf;
  end excel_encrypt;
$END
  --
  function finish( p_password varchar2 := null )
  return blob
  is
    t_excel blob;
    t_yyy blob;
    t_xxx clob;
    t_c number;
    t_h number;
    t_w number;
    t_cw number;
    s pls_integer;
    t_row_ind pls_integer;
    t_col_min pls_integer;
    t_col_max pls_integer;
    t_col_ind pls_integer;
    t_len pls_integer;
  begin
    dbms_lob.createtemporary( t_excel, true );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    s := workbook.sheets.first;
    while s is not null
    loop
      t_xxx := t_xxx || ( '
<Override PartName="/xl/worksheets/sheet' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' );
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := t_xxx || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    s := workbook.sheets.first;
    while s is not null
    loop
      if workbook.sheets( s ).comments.count > 0
      then
        t_xxx := t_xxx || ( '
<Override PartName="/xl/comments' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' );
      end if;
      if workbook.sheets( s ).drawings.count > 0
      then
        t_xxx := t_xxx || ( '
<Override ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" PartName="/xl/drawings/drawing' || s || '.xml"/>' );
      end if;
      s := workbook.sheets.next( s );
    end loop;
    if workbook.images.count > 0
    then
      t_xxx := t_xxx || '
<Default ContentType="image/png" Extension="png"/>';
    end if;
    for i in 1 .. workbook.tables.count
    loop
      t_xxx := t_xxx || ( '
<Override PartName="/xl/tables/table' || to_char(i) || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' );
    end loop;
    t_xxx := t_xxx || '
</Types>';
    add1xml( t_excel, '[Content_Types].xml', t_xxx );
    --
    declare
      l_name  varchar2(32767);
      l_ref   varchar2(100);
      l_row   tp_cells;
      l_table tp_table;
      type tp_test is table of varchar2(32767);
      l_test  tp_test;
    begin
      for i in 1 .. workbook.tables.count
      loop
        l_table := workbook.tables( i );
        l_ref := '"' || alfan_col( l_table.column_start ) || l_table.row_start ||
                 ':' || alfan_col( l_table.column_end ) || l_table.row_end || '"';
        t_xxx := ( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' ||
                   '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' ||
                   ' id="' || to_char( i ) || '"' ||
                   ' name="' || l_table.name || '"' ||
                   ' displayName="' || l_table.name || '"' ||
                   ' ref=' || l_ref ||
                   ' totalsRowShown="0">' ||
                   '<autoFilter ref=' || l_ref || '/>' ||
                   '<tableColumns count="' || to_char( 1 + l_table.column_end - l_table.column_start ) || '">'
                 );
        --
        l_test := tp_test();
        l_row := workbook.sheets( l_table.sheet ).rows( l_table.row_start );
        for j in l_table.column_start .. l_table.column_end
        loop
          l_name := workbook.str_ind( l_row( j ).value );
          if l_name member of l_test
          then
            raise_application_error( -20010, 'Table Header "' || l_name || '" appears multiple times in Table "' || l_table.name || '"' );
          end if;
          l_test := l_test multiset union tp_test( l_name );
          t_xxx := t_xxx || ( '<tableColumn id="' || to_char( j ) || '" name="' || l_name || '"/>' );
        end loop;
        --
        t_xxx := t_xxx || ( '</tableColumns>' ||
                            '<tableStyleInfo name="' || l_table.style || '"' ||
                            ' showFirstColumn="0" showLastColumn="0"' ||
                            ' showRowStripes="1" showColumnStripes="0"' ||
                            '/></table>'
                          );
        add1xml( t_excel, 'xl/tables/table' || to_char(i) || '.xml', t_xxx );
      end loop;
    end;
    --
    t_xxx := ( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context( 'userenv', 'os_user' ) || '</dc:creator>
<dc:description>Build by version:' || c_version || '</dc:description>
<cp:lastModifiedBy>' || sys_context( 'userenv', 'os_user' ) || '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:modified>
</cp:coreProperties>' );
    add1xml( t_excel, 'docProps/core.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>' || workbook.sheets.count || '</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="' || workbook.sheets.count || '" baseType="lpstr">';
    s := workbook.sheets.first;
    while s is not null
    loop
      t_xxx := t_xxx || ( '
<vt:lpstr>' || workbook.sheets( s ).name || '</vt:lpstr>' );
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := t_xxx || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
    add1xml( t_excel, 'docProps/app.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    add1xml( t_excel, '_rels/.rels', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    if workbook.numFmts.count > 0
    then
      t_xxx := t_xxx || ( '<numFmts count="' || workbook.numFmts.count || '">' );
      for n in 1 .. workbook.numFmts.count
      loop
        t_xxx := t_xxx || ( '<numFmt numFmtId="' || workbook.numFmts( n ).numFmtId || '" formatCode="' || workbook.numFmts( n ).formatCode || '"/>' );
      end loop;
      t_xxx := t_xxx || '</numFmts>';
    end if;
    t_xxx := t_xxx || ( '<fonts count="' || workbook.fonts.count || '" x14ac:knownFonts="1">' );
    for f in 0 .. workbook.fonts.count - 1
    loop
      t_xxx := t_xxx || ( '<font>' ||
        case when workbook.fonts( f ).bold then '<b/>' end ||
        case when workbook.fonts( f ).italic then '<i/>' end ||
        case when workbook.fonts( f ).underline then '<u/>' end ||
'<sz val="' || to_char( workbook.fonts( f ).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )  || '"/>' ||
    case when workbook.fonts( f ).rgb is not null
      then add_rgb( workbook.fonts( f ).rgb )
      else '<color theme="' || workbook.fonts( f ).theme || '"/>'
    end ||
'<name val="' || workbook.fonts( f ).name || '"/>
<family val="' || workbook.fonts( f ).family || '"/>
<scheme val="none"/>
</font>' );
    end loop;
    t_xxx := t_xxx || ( '</fonts>
<fills count="' || workbook.fills.count || '">' );
    for f in 0 .. workbook.fills.count - 1
    loop
      t_xxx := t_xxx || ( '<fill><patternFill patternType="' || workbook.fills( f ).patternType || '">' ||
         add_rgb( workbook.fills( f ).fgRGB, 'fgColor' ) ||
         '</patternFill></fill>' );
    end loop;
    t_xxx := t_xxx || ( '</fills>
<borders count="' || workbook.borders.count || '">' );
    for b in 0 .. workbook.borders.count - 1
    loop
      t_xxx := t_xxx || ( '<border>' ||
         case when workbook.borders( b ).left   is null then '<left/>'   else '<left style="'   || workbook.borders( b ).left   || '">' || add_rgb( workbook.borders( b ).rgb ) || '</left>' end ||
         case when workbook.borders( b ).right  is null then '<right/>'  else '<right style="'  || workbook.borders( b ).right  || '">' || add_rgb( workbook.borders( b ).rgb ) || '</right>' end ||
         case when workbook.borders( b ).top    is null then '<top/>'    else '<top style="'    || workbook.borders( b ).top    || '">' || add_rgb( workbook.borders( b ).rgb ) || '</top>' end ||
         case when workbook.borders( b ).bottom is null then '<bottom/>' else '<bottom style="' || workbook.borders( b ).bottom || '">' || add_rgb( workbook.borders( b ).rgb ) || '</bottom>' end ||
         '</border>' );
    end loop;
    t_xxx := t_xxx || ( '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || ( workbook.cellXfs.count + 1 ) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' );
    for x in 1 .. workbook.cellXfs.count
    loop
      t_xxx := t_xxx || ( '<xf numFmtId="' || workbook.cellXfs( x ).numFmtId || '" fontId="' || workbook.cellXfs( x ).fontId || '" fillId="' || workbook.cellXfs( x ).fillId || '" borderId="' || workbook.cellXfs( x ).borderId || '">' );
      if (  workbook.cellXfs( x ).alignment.horizontal is not null
         or workbook.cellXfs( x ).alignment.vertical is not null
         or workbook.cellXfs( x ).alignment.wrapText
         )
      then
        t_xxx := t_xxx || ( '<alignment' ||
          case when workbook.cellXfs( x ).alignment.horizontal is not null then ' horizontal="' || workbook.cellXfs( x ).alignment.horizontal || '"' end ||
          case when workbook.cellXfs( x ).alignment.vertical is not null then ' vertical="' || workbook.cellXfs( x ).alignment.vertical || '"' end ||
          case when workbook.cellXfs( x ).alignment.wrapText then ' wrapText="true"' end || '/>' );
      end if;
      t_xxx := t_xxx || '</xf>';
    end loop;
    t_xxx := t_xxx || ( '</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
<extLst>
<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
</ext>
</extLst>
</styleSheet>' );
    add1xml( t_excel, 'xl/styles.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
    s := workbook.sheets.first;
    while s is not null
    loop
      t_xxx := t_xxx || ( '
<sheet name="' || workbook.sheets( s ).name || '" sheetId="' || s || '" r:id="rId' || ( 9 + s ) || '"/>' );
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := t_xxx || '</sheets>';
    if workbook.defined_names.count > 0
    then
      t_xxx := t_xxx || '<definedNames>';
      for s in 1 .. workbook.defined_names.count
      loop
        t_xxx := t_xxx || ( '
<definedName name="' || workbook.defined_names( s ).name || '"' ||
            case when workbook.defined_names( s ).sheet is not null then ' localSheetId="' || to_char( workbook.defined_names( s ).sheet ) || '"' end ||
            '>' || workbook.defined_names( s ).ref || '</definedName>' );
      end loop;
      t_xxx := t_xxx || '</definedNames>';
    end if;
    t_xxx := t_xxx || '<calcPr calcId="144525"/></workbook>';
    add1xml( t_excel, 'xl/workbook.xml', t_xxx );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:shade val="51000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="80000">
<a:schemeClr val="phClr">
<a:shade val="93000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="94000"/>
<a:satMod val="135000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>';
    add1xml( t_excel, 'xl/theme/theme1.xml', t_xxx );
    s := workbook.sheets.first;
    while s is not null
    loop
      t_col_min := 16384;
      t_col_max := 1;
      t_row_ind := workbook.sheets( s ).rows.first;
      while t_row_ind is not null
      loop
        t_col_min := least( t_col_min, nvl( workbook.sheets( s ).rows( t_row_ind ).first, t_col_min ) );
        t_col_max := greatest( t_col_max, nvl( workbook.sheets( s ).rows( t_row_ind ).last, t_col_max ) );
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      end loop;
      addtxt2utf8blob_init( t_yyy );
      addtxt2utf8blob( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' ||
case when workbook.sheets( s ).tabcolor is not null then '<sheetPr>' || add_rgb( workbook.sheets( s ).tabcolor, 'tabColor' ) || '</sheetPr>' end ||
'<dimension ref="' || alfan_col( t_col_min ) || workbook.sheets( s ).rows.first || ':' || alfan_col( t_col_max ) || workbook.sheets( s ).rows.last || '"/>
<sheetViews>
<sheetView' ||
  case when workbook.sheets( s ).grid_color_idx is not null then ' defaultGridColor="0" colorId="' || to_char( workbook.sheets( s ).grid_color_idx ) || '"' end ||
  case when not workbook.sheets( s ).show_gridlines then ' showGridLines="0"' end ||
  case when not workbook.sheets( s ).show_headers then ' showRowColHeaders="0"' end ||
  case when s = 1 then ' tabSelected="1"' end || ' workbookViewId="0">'
                     , t_yyy
                     );
      if workbook.sheets( s ).freeze_rows > 0 and workbook.sheets( s ).freeze_cols > 0
      then
        addtxt2utf8blob( '<pane xSplit="' || workbook.sheets( s ).freeze_cols || '" '
                          || 'ySplit="' || workbook.sheets( s ).freeze_rows || '" '
                          || 'topLeftCell="' || alfan_col( workbook.sheets( s ).freeze_cols + 1 ) || ( workbook.sheets( s ).freeze_rows + 1 ) || '" '
                          || 'activePane="bottomLeft" state="frozen"/>'
                       , t_yyy
                       );
      else
        if workbook.sheets( s ).freeze_rows > 0
        then
          addtxt2utf8blob( '<pane ySplit="' || workbook.sheets( s ).freeze_rows || '" topLeftCell="A' || ( workbook.sheets( s ).freeze_rows + 1 ) || '" activePane="bottomLeft" state="frozen"/>'
                         , t_yyy
                         );
        end if;
        if workbook.sheets( s ).freeze_cols > 0
        then
          addtxt2utf8blob( '<pane xSplit="' || workbook.sheets( s ).freeze_cols || '" topLeftCell="' || alfan_col( workbook.sheets( s ).freeze_cols + 1 ) || '1" activePane="bottomLeft" state="frozen"/>'
                         , t_yyy
                         );
        end if;
      end if;
      addtxt2utf8blob( '</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
                     , t_yyy
                     );
      if workbook.sheets( s ).widths.count > 0
      then
        addtxt2utf8blob( '<cols>', t_yyy );
        t_col_ind := workbook.sheets( s ).widths.first;
        while t_col_ind is not null
        loop
          addtxt2utf8blob( '<col min="' || t_col_ind || '" max="' || t_col_ind || '" width="' || to_char( workbook.sheets( s ).widths( t_col_ind ), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" customWidth="1"/>', t_yyy );
          t_col_ind := workbook.sheets( s ).widths.next( t_col_ind );
        end loop;
        addtxt2utf8blob( '</cols>', t_yyy );
      end if;
      addtxt2utf8blob( '<sheetData>', t_yyy );
      t_row_ind := workbook.sheets( s ).rows.first;
      while t_row_ind is not null
      loop
        if workbook.sheets( s ).row_fmts.exists( t_row_ind ) and workbook.sheets( s ).row_fmts( t_row_ind ).height is not null
        then
          addtxt2utf8blob( '<row r="' || t_row_ind || '" spans="' || t_col_min || ':' || t_col_max || '" customHeight="1" ht="'
                         || to_char( workbook.sheets( s ).row_fmts( t_row_ind ).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" >', t_yyy );
        else
          addtxt2utf8blob( '<row r="' || t_row_ind || '" spans="' || t_col_min || ':' || t_col_max || '">', t_yyy );
        end if;
        t_col_ind := workbook.sheets( s ).rows( t_row_ind ).first;
        while t_col_ind is not null
        loop
          addtxt2utf8blob( '<c r="' || alfan_col( t_col_ind ) || t_row_ind || '"'
                 || ' ' || workbook.sheets( s ).rows( t_row_ind )( t_col_ind ).style
                 || '><v>'
                 || to_char( workbook.sheets( s ).rows( t_row_ind )( t_col_ind ).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )
                 || '</v></c>', t_yyy );
          t_col_ind := workbook.sheets( s ).rows( t_row_ind ).next( t_col_ind );
        end loop;
        addtxt2utf8blob( '</row>', t_yyy );
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      end loop;
      addtxt2utf8blob( '</sheetData>', t_yyy );
      for a in 1 ..  workbook.sheets( s ).autofilters.count
      loop
        addtxt2utf8blob( '<autoFilter ref="' ||
            alfan_col( nvl( workbook.sheets( s ).autofilters( a ).column_start, t_col_min ) ) ||
            nvl( workbook.sheets( s ).autofilters( a ).row_start, workbook.sheets( s ).rows.first ) || ':' ||
            alfan_col( coalesce( workbook.sheets( s ).autofilters( a ).column_end, workbook.sheets( s ).autofilters( a ).column_start, t_col_max ) ) ||
            nvl( workbook.sheets( s ).autofilters( a ).row_end, workbook.sheets( s ).rows.last ) || '"/>', t_yyy );
      end loop;
      if workbook.sheets( s ).mergecells.count > 0
      then
        addtxt2utf8blob( '<mergeCells count="' || to_char( workbook.sheets( s ).mergecells.count ) || '">', t_yyy );
        for m in 1 ..  workbook.sheets( s ).mergecells.count
        loop
          addtxt2utf8blob( '<mergeCell ref="' || workbook.sheets( s ).mergecells( m ) || '"/>', t_yyy );
        end loop;
        addtxt2utf8blob( '</mergeCells>', t_yyy );
      end if;
--
      if workbook.sheets( s ).validations.count > 0
      then
        addtxt2utf8blob( '<dataValidations count="' || to_char( workbook.sheets( s ).validations.count ) || '">', t_yyy );
        for m in 1 ..  workbook.sheets( s ).validations.count
        loop
          addtxt2utf8blob( '<dataValidation' ||
              ' type="' || workbook.sheets( s ).validations( m ).type || '"' ||
              ' errorStyle="' || workbook.sheets( s ).validations( m ).errorstyle || '"' ||
              ' allowBlank="' || case when nvl( workbook.sheets( s ).validations( m ).allowBlank, true ) then '1' else '0' end || '"' ||
              ' sqref="' || workbook.sheets( s ).validations( m ).sqref || '"', t_yyy );
          if workbook.sheets( s ).validations( m ).prompt is not null
          then
            addtxt2utf8blob( ' showInputMessage="1" prompt="' || workbook.sheets( s ).validations( m ).prompt || '"', t_yyy );
            if workbook.sheets( s ).validations( m ).title is not null
            then
              addtxt2utf8blob( ' promptTitle="' || workbook.sheets( s ).validations( m ).title || '"', t_yyy );
            end if;
          end if;
          if workbook.sheets( s ).validations( m ).showerrormessage
          then
            addtxt2utf8blob( ' showErrorMessage="1"', t_yyy );
            if workbook.sheets( s ).validations( m ).error_title is not null
            then
              addtxt2utf8blob( ' errorTitle="' || workbook.sheets( s ).validations( m ).error_title || '"', t_yyy );
            end if;
            if workbook.sheets( s ).validations( m ).error_txt is not null
            then
              addtxt2utf8blob( ' error="' || workbook.sheets( s ).validations( m ).error_txt || '"', t_yyy );
            end if;
          end if;
          addtxt2utf8blob( '>', t_yyy );
          if workbook.sheets( s ).validations( m ).formula1 is not null
          then
            addtxt2utf8blob( '<formula1>' || workbook.sheets( s ).validations( m ).formula1 || '</formula1>', t_yyy );
          end if;
          if workbook.sheets( s ).validations( m ).formula2 is not null
          then
            addtxt2utf8blob( '<formula2>' || workbook.sheets( s ).validations( m ).formula2 || '</formula2>', t_yyy );
          end if;
          addtxt2utf8blob( '</dataValidation>', t_yyy );
        end loop;
        addtxt2utf8blob( '</dataValidations>', t_yyy );
      end if;
--
      if workbook.sheets( s ).hyperlinks.count > 0
      then
        addtxt2utf8blob( '<hyperlinks>', t_yyy );
        for h in 1 ..  workbook.sheets( s ).hyperlinks.count
        loop
          addtxt2utf8blob( '<hyperlink ref="' || workbook.sheets( s ).hyperlinks( h ).cell || '"', t_yyy );
          if workbook.sheets( s ).hyperlinks( h ).url is null
          then
            addtxt2utf8blob( ' location="' || workbook.sheets( s ).hyperlinks( h ).location || '"', t_yyy );
          else
            addtxt2utf8blob( ' r:id="rId' || ( 3 + h ) || '"', t_yyy );
          end if;
          if workbook.sheets( s ).hyperlinks( h ).tooltip is not null
          then
            addtxt2utf8blob( ' tooltip="' || workbook.sheets( s ).hyperlinks( h ).tooltip || '"/>', t_yyy );
          else
            addtxt2utf8blob( '/>', t_yyy );
          end if;
        end loop;
        addtxt2utf8blob( '</hyperlinks>', t_yyy );
      end if;
      addtxt2utf8blob( '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', t_yyy );
      if workbook.sheets( s ).drawings.count > 0
      then
        addtxt2utf8blob( '<drawing r:id="rId3"/>', t_yyy );
      end if;
      if workbook.sheets( s ).comments.count > 0
      then
        addtxt2utf8blob( '<legacyDrawing r:id="rId1"/>', t_yyy );
      end if;
      --
      declare
        l_cnt pls_integer := 0;
      begin
        for i in 1 .. workbook.tables.count
        loop
          if workbook.tables( i ).sheet = s
          then
            l_cnt := l_cnt + 1;
          end if;
        end loop;
        if l_cnt > 0
        then
          addtxt2utf8blob( '<tableParts count="' || to_char( l_cnt ) || '">', t_yyy) ;
          for i in 1 .. workbook.tables.count
          loop
            if workbook.tables( i ).sheet = s
            then
              addtxt2utf8blob( '<tablePart r:id="rId' || to_char( 10000 + i ) || '"/>', t_yyy );
            end if;
          end loop;
          addtxt2utf8blob( '</tableParts>', t_yyy );
        end if;
      end;
      --
      addtxt2utf8blob( '</worksheet>', t_yyy );
      addtxt2utf8blob_finish( t_yyy );
      add1file( t_excel, 'xl/worksheets/sheet' || s || '.xml', t_yyy );
      --
      if workbook.images.count > 0
      then
        t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        for i in 1 .. workbook.images.count
        loop
          t_xxx := t_xxx || ( '<Relationship Id="rId'
                         || i || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' || 'image' || i || '.png'
                         || '"/>' );
        end loop;
        t_xxx := t_xxx || '</Relationships>';
        add1xml( t_excel, 'xl/drawings/_rels/drawing' || s || '.xml.rels', t_xxx );
      end if;
      --
      if (  workbook.sheets( s ).hyperlinks.count > 0
         or workbook.sheets( s ).comments.count > 0
         or workbook.sheets( s ).drawings.count > 0
         or workbook.tables.count > 0
         )
      then
        t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        if workbook.sheets( s ).comments.count > 0
        then
          t_xxx := t_xxx || ( '<Relationship Id="rId2" Target="../comments' || s || '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"/>' );
          t_xxx := t_xxx || ( '<Relationship Id="rId1" Target="../drawings/vmlDrawing' || s || '.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>' );
        end if;
        if workbook.sheets( s ).drawings.count > 0
        then
          t_xxx := t_xxx || ( '<Relationship Id="rId3" Target="../drawings/drawing' || s || '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>' );
        end if;
        for h in 1 ..  workbook.sheets( s ).hyperlinks.count
        loop
          if workbook.sheets( s ).hyperlinks( h ).url is not null
          then
            t_xxx := t_xxx || ( '<Relationship Id="rId' || ( 3 + h ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets( s ).hyperlinks( h ).url || '" TargetMode="External"/>' );
          end if;
        end loop;
        for i in 1 .. workbook.tables.count
        loop
          if workbook.tables( i ).sheet = s
          then
            t_xxx := t_xxx ||  ( '<Relationship Id="rId' || to_char( 10000 + i ) || '"' ||
                                 ' Target="../tables/table' || to_char( i ) || '.xml"' ||
                                 ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"/>' );
          end if;
        end loop;
        t_xxx := t_xxx || '</Relationships>';
        add1xml( t_excel, 'xl/worksheets/_rels/sheet' || s || '.xml.rels', t_xxx );
        --
        if workbook.sheets( s ).drawings.count > 0
        then
          t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
          for i in 1 .. workbook.sheets( s ).drawings.count
          loop
            t_xxx := t_xxx || finish_drawing( workbook.sheets( s ).drawings( i ), i, s );
          end loop;
          t_xxx := t_xxx || '</xdr:wsDr>';
          add1xml( t_excel, 'xl/drawings/drawing' || s || '.xml', t_xxx );
        end if;
--
        if workbook.sheets( s ).comments.count > 0
        then
          declare
            cnt pls_integer;
            author_ind tp_author;
          begin
            authors.delete;
            for c in 1 .. workbook.sheets( s ).comments.count
            loop
              authors( nvl( workbook.sheets( s ).comments( c ).author, ' ' ) ) := 0;
            end loop;
            t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
            cnt := 0;
            author_ind := authors.first;
            while author_ind is not null or authors.next( author_ind ) is not null
            loop
              authors( author_ind ) := cnt;
              t_xxx := t_xxx || ( '<author>' || author_ind || '</author>' );
              cnt := cnt + 1;
              author_ind := authors.next( author_ind );
            end loop;
          end;
          t_xxx := t_xxx || '</authors><commentList>';
          for c in 1 .. workbook.sheets( s ).comments.count
          loop
            t_xxx := t_xxx || ( '<comment ref="' || alfan_col( workbook.sheets( s ).comments( c ).column ) ||
               to_char( workbook.sheets( s ).comments( c ).row ) || '"' ||
               ' authorId="' || authors( nvl( workbook.sheets( s ).comments( c ).author, ' ' ) ) || '"><text>' );
            if workbook.sheets( s ).comments( c ).author is not null
            then
              t_xxx := t_xxx || ( '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
                 workbook.sheets( s ).comments( c ).author || ':</t></r>' );
            end if;
            t_xxx := t_xxx || ( '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
               case when workbook.sheets( s ).comments( c ).author is not null then '
' end || workbook.sheets( s ).comments( c ).text || '</t></r></text></comment>' );
          end loop;
          t_xxx := t_xxx || '</commentList></comments>';
          add1xml( t_excel, 'xl/comments' || s || '.xml', t_xxx );
          t_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
          for c in 1 .. workbook.sheets( s ).comments.count
          loop
            t_xxx := t_xxx || ( '<v:shape id="_x0000_s' || to_char( c ) || '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' || to_char( c ) || ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' );
            t_w := workbook.sheets( s ).comments( c ).width;
            t_c := 1;
            loop
              if workbook.sheets( s ).widths.exists( workbook.sheets( s ).comments( c ).column + t_c )
              then
                t_cw := 256 * workbook.sheets( s ).widths( workbook.sheets( s ).comments( c ).column + t_c );
                t_cw := trunc( ( t_cw + 18 ) / 256 * 7); -- assume default 11 point Calibri
              else
                t_cw := 64;
              end if;
              exit when t_w < t_cw;
              t_c := t_c + 1;
              t_w := t_w - t_cw;
            end loop;
            t_h := workbook.sheets( s ).comments( c ).height;
            t_xxx := t_xxx || ( '<x:Anchor>' || workbook.sheets( s ).comments( c ).column || ',15,' ||
                       workbook.sheets( s ).comments( c ).row || ',30,' ||
                       ( workbook.sheets( s ).comments( c ).column + t_c - 1 ) || ',' || round( t_w ) || ',' ||
                       ( workbook.sheets( s ).comments( c ).row + 1 + trunc( t_h / 20 ) ) || ',' || mod( t_h, 20 ) || '</x:Anchor>' );
            t_xxx := t_xxx || ( '<x:AutoFill>False</x:AutoFill><x:Row>' ||
              ( workbook.sheets( s ).comments( c ).row - 1 ) || '</x:Row><x:Column>' ||
              ( workbook.sheets( s ).comments( c ).column - 1 ) || '</x:Column></x:ClientData></v:shape>' );
          end loop;
          t_xxx := t_xxx || '</xml>';
          add1xml( t_excel, 'xl/drawings/vmlDrawing' || s || '.vml', t_xxx );
        end if;
      end if;
--
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    s := workbook.sheets.first;
    while s is not null
    loop
      t_xxx := t_xxx || ( '
<Relationship Id="rId' || ( 9 + s ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s || '.xml"/>' );
      s := workbook.sheets.next( s );
    end loop;
    t_xxx := t_xxx || '</Relationships>';
    add1xml( t_excel, 'xl/_rels/workbook.xml.rels', t_xxx );
    --
    for i in 1 .. workbook.images.count
    loop
      add1file( t_excel, 'xl/media/image' || i || '.png', workbook.images(i).img );
    end loop;
    --
    addtxt2utf8blob_init( t_yyy );
    addtxt2utf8blob( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || workbook.str_cnt || '" uniqueCount="' || workbook.strings.count || '">'
                  , t_yyy
                  );
    for i in 0 .. workbook.str_ind.count - 1
    loop
      addtxt2utf8blob( '<si><t xml:space="preserve">' || dbms_xmlgen.convert( substr( workbook.str_ind( i ), 1, 32000 ) ) || '</t></si>', t_yyy );
    end loop;
    addtxt2utf8blob( '</sst>', t_yyy );
    addtxt2utf8blob_finish( t_yyy );
    add1file( t_excel, 'xl/sharedStrings.xml', t_yyy );
    finish_zip( t_excel );
    clear_workbook;
$IF as_xlsx.use_dbms_crypto
$THEN
    if p_password is not null
    then
      return excel_encrypt( t_excel, p_password );
    else
      return t_excel;
    end if;
$ELSE
    return t_excel;
$END
  end finish;
--
  procedure save
    ( p_directory varchar2
    , p_filename varchar2
    , p_password varchar2 := null
    )
  is
  begin
    blob2file( finish( p_password ), p_directory, p_filename );
  end;
--
  function query2sheet
    ( p_c              in out integer
    , p_column_headers boolean
    , p_directory      varchar2
    , p_filename       varchar2
    , p_sheet          pls_integer
    , p_UseXf          boolean
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2
    , p_title_xfid     pls_integer
    , p_col            pls_integer
    , p_row            pls_integer
    , p_autofilter     boolean
    , p_table_style    varchar2
    )
  return number
  is
    l_null varchar2(1);
    t_sheet pls_integer;
    t_col_cnt integer;
    t_desc_tab dbms_sql.desc_tab2;
    d_tab dbms_sql.date_table;
    n_tab dbms_sql.number_table;
    v_tab dbms_sql.varchar2_table;
    t_bulk_size pls_integer := 200;
    t_r integer;
    t_col     pls_integer;
    t_cur_row pls_integer;
    t_useXf boolean := g_useXf;
    type tp_XfIds is table of varchar2(50) index by pls_integer;
    t_XfIds tp_XfIds;
    t_null_number number;
    l_rows number;
    l_horizontal varchar2(3999);
    l_right       boolean;
    l_center_cont boolean;
  begin
    if p_sheet is null
    then
      new_sheet;
    end if;
    t_sheet := coalesce( p_sheet, workbook.sheets.count );
    setUseXf( true );
    t_col := coalesce( p_col, 1 ) - 1;
    t_cur_row := coalesce( p_row, 1 );
    if p_title is not null
    then
      t_cur_row := t_cur_row + 1;
      if p_title_xfid is not null and workbook.cellXfs.exists( p_title_xfid )
      then
        l_horizontal := lower( workbook.cellXfs( p_title_xfid ).alignment.horizontal );
        l_right := l_horizontal = 'right';
        l_center_cont := l_horizontal = 'centercontinuous';
      end if;
      l_right := nvl( l_right, false );
      l_center_cont := nvl( l_center_cont, false );
    end if;
    dbms_sql.describe_columns2( p_c, t_col_cnt, t_desc_tab );
    for c in 1 .. t_col_cnt
    loop
      if p_title is not null
      then
        if ( c = 1 and not l_right ) or ( c = t_col_cnt and l_right )
        then
          cell( t_col + c, t_cur_row - 1, p_title, p_sheet => t_sheet );
          if p_title_xfid is not null
          then
            workbook.sheets( t_sheet ).rows( t_cur_row - 1 )( t_col + c ).style := 't="s" s="' || p_title_xfid || '"';
          end if;
        elsif l_center_cont
        then
          cell( t_col + c, t_cur_row - 1, t_null_number, p_sheet => t_sheet );
          workbook.sheets( t_sheet ).rows( t_cur_row - 1 )( t_col + c ).style := 's="' || p_title_xfid || '"';
        end if;
      end if;
      if p_column_headers
      then
        cell( t_col + c, t_cur_row, t_desc_tab( c ).col_name, p_sheet => t_sheet );
      end if;
      case
        when t_desc_tab( c ).col_type in ( 2, 100, 101 )
        then
          dbms_sql.define_array( p_c, c, n_tab, t_bulk_size, 1 );
        when t_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181, 231 )
        then
          dbms_sql.define_array( p_c, c, d_tab, t_bulk_size, 1 );
          t_XfIds(c) := get_XfId( t_sheet, t_col + c, t_cur_row, get_numFmt( coalesce( p_date_format, 'dd/mm/yyyy' ) ), null, null );
        when t_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )
        then
          dbms_sql.define_array( p_c, c, v_tab, t_bulk_size, 1 );
        else
          null;
      end case;
    end loop;
    --
    setUseXf( p_UseXf );
    if p_column_headers
    then
      t_cur_row := t_cur_row + 1;
    end if;
    --
    l_rows := 0;
    loop
      t_r := dbms_sql.fetch_rows( p_c );
      if t_r > 0
      then
        for c in 1 .. t_col_cnt
        loop
          case
            when t_desc_tab( c ).col_type in ( 2, 100, 101 )
            then
              dbms_sql.column_value( p_c, c, n_tab );
              for i in 0 .. t_r - 1
              loop
                if n_tab( i + n_tab.first ) is not null
                then
                  cell( t_col + c, t_cur_row + i, n_tab( i + n_tab.first ), p_sheet => t_sheet );
                else
                  cell( t_col + c, t_cur_row + i, t_null_number, p_sheet => t_sheet );
                end if;
              end loop;
              n_tab.delete;
            when t_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181, 231 )
            then
              dbms_sql.column_value( p_c, c, d_tab );
              for i in 0 .. t_r - 1
              loop
                if d_tab( i + d_tab.first ) is not null
                then
                  if g_useXf
                  then
                    cell( t_col + c, t_cur_row + i, d_tab( i + d_tab.first ), p_sheet => t_sheet );
                  else
                    query_date_cell( t_col + c, t_cur_row + i, d_tab( i + d_tab.first ), t_sheet, t_XfIds(c) );
                  end if;
                else
                  cell( t_col + c, t_cur_row + i, t_null_number, p_sheet => t_sheet );
                end if;
              end loop;
              d_tab.delete;
            when t_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )
            then
              dbms_sql.column_value( p_c, c, v_tab );
              for i in 0 .. t_r - 1
              loop
                if v_tab( i + v_tab.first ) is not null
                then
                  cell( t_col + c, t_cur_row + i, v_tab( i + v_tab.first ), p_sheet => t_sheet );
                else
                  cell( t_col + c, t_cur_row + i, t_null_number, p_sheet => t_sheet );
                end if;
              end loop;
              v_tab.delete;
            else
              null;
          end case;
        end loop;
      end if;
      l_rows := l_rows + t_r;
      t_cur_row := t_cur_row + t_r;
      exit when t_r != t_bulk_size;
    end loop;
    dbms_sql.close_cursor( p_c );
    --
    if p_autofilter and p_table_style is null and p_column_headers
    then
      set_autofilter
        ( p_column_start => t_col + 1
        , p_column_end   => t_col + t_col_cnt
        , p_row_start    => coalesce( p_row, 1 ) + case when p_title is null then 0 else 1 end
        , p_row_end      => t_cur_row - 1
        , p_sheet        => t_sheet
        );
    end if;
    --
    if p_table_style is not null and p_column_headers
    then
      set_table
        ( p_column_start => t_col + 1
        , p_column_end   => t_col + t_col_cnt
        , p_row_start    => coalesce( p_row, 1 ) + case when p_title is null then 0 else 1 end
        , p_row_end      => t_cur_row - 1
        , p_style        => p_table_style
        , p_sheet        => t_sheet
        );
    end if;
    --
    if ( p_directory is not null and  p_filename is not null )
    then
      save( p_directory, p_filename );
    end if;
    --
    setUseXf( t_useXf );
    return l_rows;
  exception
    when others
    then
      if dbms_sql.is_open( p_c )
      then
        dbms_sql.close_cursor( p_c );
      end if;
      setUseXf( t_useXf );
      return null;
  end query2sheet;
  --
  function query2sheet
    ( p_sql            varchar2
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  return number
  is
    t_c integer;
    t_r integer;
  begin
    t_c := dbms_sql.open_cursor;
    dbms_sql.parse( t_c, p_sql, dbms_sql.native );
    t_r := dbms_sql.execute( t_c );
    return query2sheet
             ( p_c              => t_c
             , p_column_headers => p_column_headers
             , p_directory      => p_directory
             , p_filename       => p_filename
             , p_sheet          => p_sheet
             , p_UseXf          => p_UseXf
             , p_date_format    => p_date_format
             , p_title          => p_title
             , p_title_xfid     => p_title_xfid
             , p_col            => p_col
             , p_row            => p_row
             , p_autofilter     => p_autofilter
             , p_table_style    => p_table_style
             );
  end;
  --
  function query2sheet
    ( p_rc             in out sys_refcursor
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  return number
  is
    t_c integer;
    t_r integer;
  begin
    t_c := dbms_sql.to_cursor_number( p_rc );
    return query2sheet
             ( p_c              => t_c
             , p_column_headers => p_column_headers
             , p_directory      => p_directory
             , p_filename       => p_filename
             , p_sheet          => p_sheet
             , p_UseXf          => p_UseXf
             , p_date_format    => p_date_format
             , p_title          => p_title
             , p_title_xfid     => p_title_xfid
             , p_col            => p_col
             , p_row            => p_row
             , p_autofilter     => p_autofilter
             , p_table_style    => p_table_style
             );
  end;
  --
  procedure query2sheet
    ( p_sql            varchar2
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  is
    l_dummy number;
  begin
    l_dummy := query2sheet
                 ( p_sql            => p_sql
                 , p_column_headers => p_column_headers
                 , p_directory      => p_directory
                 , p_filename       => p_filename
                 , p_sheet          => p_sheet
                 , p_UseXf          => p_UseXf
                 , p_date_format    => p_date_format
                 , p_title          => p_title
                 , p_title_xfid     => p_title_xfid
                 , p_col            => p_col
                 , p_row            => p_row
                 , p_autofilter     => p_autofilter
                 , p_table_style    => p_table_style
                 );
  end;
--
  procedure query2sheet
    ( p_rc             in out sys_refcursor
    , p_column_headers boolean     := true
    , p_directory      varchar2    := null
    , p_filename       varchar2    := null
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  is
    l_dummy number;
  begin
    l_dummy := query2sheet
                 ( p_rc             => p_rc
                 , p_column_headers => p_column_headers
                 , p_directory      => p_directory
                 , p_filename       => p_filename
                 , p_sheet          => p_sheet
                 , p_UseXf          => p_UseXf
                 , p_date_format    => p_date_format
                 , p_title          => p_title
                 , p_title_xfid     => p_title_xfid
                 , p_col            => p_col
                 , p_row            => p_row
                 , p_autofilter     => p_autofilter
                 , p_table_style    => p_table_style
                 );
  end;
  --
  procedure setUseXf( p_val boolean := true )
  is
  begin
    g_useXf := p_val;
  end;
--
  procedure add_image
    ( p_col pls_integer
    , p_row pls_integer
    , p_img blob
    , p_name varchar2 := ''
    , p_title varchar2 := ''
    , p_description varchar2 := ''
    , p_scale number := null
    , p_sheet pls_integer := null
    , p_width pls_integer := null
    , p_height pls_integer := null
    )
  is
    l_hash raw(4);
    l_image tp_image;
    l_idx pls_integer;
    l_sheet pls_integer := coalesce( p_sheet, workbook.sheets.count );
    l_drawing tp_drawing;
    l_ind number;
    l_len number;
    l_buf raw(32);
    l_hex varchar2(8);
  begin
    l_hash := utl_raw.substr( utl_compress.lz_compress( p_img ), -8, 4 );
    for i in 1 .. workbook.images.count
    loop
      if workbook.images(i).hash = l_hash
      then
        l_idx := i;
        exit;
      end if;
    end loop;
    if l_idx is null
    then
      l_idx := workbook.images.count + 1;
      dbms_lob.createtemporary( l_image.img, true );
      dbms_lob.copy( l_image.img, p_img, dbms_lob.lobmaxsize, 1, 1 );
      l_image.hash := l_hash;
      --
      l_buf := dbms_lob.substr( p_img, 32, 1 );
      if utl_raw.substr( l_buf, 1, 8 ) = hextoraw( '89504E470D0A1A0A' )
      then -- png
        l_ind := 9;
        loop
          l_len := to_number( dbms_lob.substr( p_img, 4, l_ind ), 'xxxxxxxx' );  -- length
          exit when l_len is null or l_ind > dbms_lob.getlength( p_img );
          case rawtohex( dbms_lob.substr( p_img, 4, l_ind + 4 ) ) -- Chunk type
            when '49484452' -- IHDR
            then
              l_image.width  := to_number( dbms_lob.substr( p_img, 4, l_ind + 8 ), 'xxxxxxxx' );
              l_image.height := to_number( dbms_lob.substr( p_img, 4, l_ind + 12 ), 'xxxxxxxx' );
              exit;
            when '49454E44' -- IEND
            then
              exit;
            else
              null;
          end case;
          l_ind := l_ind + 4 + 4 + l_len + 4;  -- Length + Chunk type + Chunk data + CRC
        end loop;
      elsif utl_raw.substr( l_buf, 1, 3 ) = hextoraw( '474946' )
      then -- gif
        l_ind := 14;
        l_buf := utl_raw.substr( l_buf, 11, 1 );
        if utl_raw.bit_and( '80', l_buf ) = '80'
        then
          l_len := to_number( utl_raw.bit_and( '07', l_buf ), 'XX' );
          l_ind := l_ind + 3 * power( 2, l_len + 1 );
        end if;
        loop
          case rawtohex( dbms_lob.substr( p_img, 1, l_ind ) )
            when '21' -- extension
            then
              l_ind := l_ind + 2; -- skip sentinel + label
              loop
                l_len := to_number( dbms_lob.substr( p_img, 1, l_ind ), 'XX' ); -- Block Size
                exit when l_len = 0;
                l_ind := l_ind + 1 + l_len; -- skip Block Size + Data Sub-block
              end loop;
              l_ind := l_ind + 1;           -- skip last Block Size
            when '2C' -- image
            then
              l_buf := dbms_lob.substr( p_img, 4, l_ind + 5 );
              l_image.width  := utl_raw.cast_to_binary_integer( utl_raw.substr( l_buf, 1, 2 ), utl_raw.little_endian );
              l_image.height := utl_raw.cast_to_binary_integer( utl_raw.substr( l_buf, 3, 2 ), utl_raw.little_endian );
              exit;
            else
              exit;
          end case;
        end loop;
      elsif utl_raw.substr( l_buf, 1, 2 ) = hextoraw( 'FFD8' ) -- SOI Start of Image
        and rawtohex( utl_raw.substr( l_buf, 3, 2 ) ) in ( 'FFE0' -- a APP0 jpg
                                                         , 'FFE1' -- a APP1 jpg
                                                         )
      then -- jpg
        l_ind := 5 + to_number( utl_raw.substr( l_buf, 5, 2 ), 'xxxx' );
        loop
          l_buf := dbms_lob.substr( p_img, 4, l_ind );
          l_hex := substr( rawtohex( l_buf ), 1, 4 );
          exit when l_hex in ( 'FFDA' -- SOS Start of Scan
                             , 'FFD9' -- EOI End Of Image
                             )
                 or substr( l_hex, 1, 2 ) != 'FF';
          if l_hex in ( 'FFD0', 'FFD1', 'FFD2', 'FFD3', 'FFD4', 'FFD5', 'FFD6', 'FFD7' -- RSTn
                      , 'FF01'  -- TEM
                      )
          then
            l_ind := l_ind + 2;
          else
            if l_hex = 'FFC0' -- SOF0 (Start Of Frame 0) marker
            then
              l_hex := rawtohex( dbms_lob.substr( p_img, 4, l_ind + 5 ) );
              l_image.width  := to_number( substr( l_hex, 5 ), 'xxxx' );
              l_image.height := to_number( substr( l_hex, 1, 4 ), 'xxxx' );
              exit;
            end if;
            l_ind := l_ind + 2 + to_number( utl_raw.substr( l_buf, 3, 2 ), 'xxxx' );
          end if;
        end loop;
      elsif utl_raw.substr( l_buf, 1, 2 ) = '424D' -- BM
      then -- bmp
        l_image.width  := to_number( utl_raw.reverse( utl_raw.substr( l_buf, 19, 4 ) ), 'XXXXXXXX' );
        l_image.height := to_number( utl_raw.reverse( utl_raw.substr( l_buf, 23, 4 ) ), 'XXXXXXXX' );
      else
        l_image.width  := nvl( p_width, 0 );
        l_image.height := nvl( p_height, 0 );
      end if;
      --
      workbook.images( l_idx ) := l_image;
    end if;
    --
    l_drawing.img_id := l_idx;
    l_drawing.row := p_row;
    l_drawing.col := p_col;
    l_drawing.scale := p_scale;
    l_drawing.name := p_name;
    l_drawing.title := p_title;
    l_drawing.description := p_description;
    workbook.sheets( l_sheet ).drawings( workbook.sheets( l_sheet ).drawings.count + 1 ) := l_drawing;
  end add_image;
--
  function get_encoding( p_encoding varchar2 := null )
  return varchar2
  is
    l_encoding varchar2(32767);
  begin
    if p_encoding is not null
    then
      if nls_charset_id( p_encoding ) is null
      then
        l_encoding := utl_i18n.map_charset( p_encoding, utl_i18n.GENERIC_CONTEXT, utl_i18n.IANA_TO_ORACLE );
      else
        l_encoding := p_encoding;
      end if;
    end if;
    return coalesce( l_encoding, 'US8PC437' ); -- IBM codepage 437
  end;
  --
  function char2raw( p_txt varchar2 character set any_cs, p_encoding varchar2 := null )
  return raw
  is
  begin
    if isnchar( p_txt )
    then -- on my 12.1 database, which is not AL32UTF8,
         -- utl_i18n.string_to_raw( p_txt, get_encoding( p_encoding ) does not work
      return utl_raw.convert( utl_i18n.string_to_raw( p_txt )
                            , get_encoding( p_encoding )
                            , nls_charset_name( nls_charset_id( 'NCHAR_CS' ) )
                            );
    end if;
    return utl_i18n.string_to_raw( p_txt, get_encoding( p_encoding ) );
  end;
  --
  procedure get_zip_info( p_zip blob, p_info out tp_zip_info )
  is
    l_ind integer;
    l_buf_sz pls_integer := 2024;
    l_start_buf integer;
    l_buf raw(32767);
  begin
    p_info.len := nvl( dbms_lob.getlength( p_zip ), 0 );
    if p_info.len < 22
    then -- no (zip) file or empty zip file
      return;
    end if;
    l_start_buf := greatest( p_info.len - l_buf_sz + 1, 1 );
    l_buf := dbms_lob.substr( p_zip, l_buf_sz, l_start_buf );
    l_ind := utl_raw.length( l_buf ) - 21;
    loop
      exit when l_ind < 1 or utl_raw.substr( l_buf, l_ind, 4 ) = c_END_OF_CENTRAL_DIRECTORY;
      l_ind := l_ind - 1;
    end loop;
    if l_ind > 0
    then
      l_ind := l_ind + l_start_buf - 1;
    else
      l_ind := p_info.len - 21;
      loop
        exit when l_ind < 1 or dbms_lob.substr( p_zip, 4, l_ind ) = c_END_OF_CENTRAL_DIRECTORY;
        l_ind := l_ind - 1;
      end loop;
    end if;
    if l_ind <= 0
    then
      raise_application_error( -20001, 'Error parsing the zipfile' );
    end if;
    l_buf := dbms_lob.substr( p_zip, 22, l_ind );
    if    utl_raw.substr( l_buf, 5, 2 ) != utl_raw.substr( l_buf, 7, 2 )  -- this disk = disk with start of Central Dir
       or utl_raw.substr( l_buf, 9, 2 ) != utl_raw.substr( l_buf, 11, 2 ) -- complete CD on this disk
    then
      raise_application_error( -20003, 'Error parsing the zipfile' );
    end if;
    p_info.idx_eocd := l_ind;
    p_info.idx_cd := little_endian( l_buf, 17, 4 ) + 1;
    p_info.cnt := little_endian( l_buf, 9, 2 );
    p_info.len_cd := p_info.idx_eocd - p_info.idx_cd;
  end;
  --
  function parse_central_file_header( p_zip blob, p_ind integer, p_cfh out tp_cfh )
  return boolean
  is
    l_tmp pls_integer;
    l_len pls_integer;
    l_buf raw(32767);
  begin
    l_buf := dbms_lob.substr( p_zip, 46, p_ind );
    if utl_raw.substr( l_buf, 1, 4 ) != c_CENTRAL_FILE_HEADER
    then
      return false;
    end if;
    p_cfh.crc32 := utl_raw.substr( l_buf, 17, 4 );
    p_cfh.n := little_endian( l_buf, 29, 2 );
    p_cfh.m := little_endian( l_buf, 31, 2 );
    p_cfh.k := little_endian( l_buf, 33, 2 );
    p_cfh.len := 46 + p_cfh.n + p_cfh.m + p_cfh.k;
    --
    p_cfh.utf8 := bitand( to_number( utl_raw.substr( l_buf, 10, 1 ), 'XX' ), 8 ) > 0;
    if p_cfh.n > 0
    then
      p_cfh.name1 := dbms_lob.substr( p_zip, least( p_cfh.n, 32767 ), p_ind + 46 );
    end if;
    --
    p_cfh.compressed_len := little_endian( l_buf, 21, 4 );
    p_cfh.original_len := little_endian( l_buf, 25, 4 );
    p_cfh.offset := little_endian( l_buf, 43, 4 );
    --
    return true;
  end;
  --
  function get_central_file_header
    ( p_zip      blob
    , p_name     varchar2 character set any_cs
    , p_idx      number
    , p_encoding varchar2
    , p_cfh      out tp_cfh
    )
  return boolean
  is
    l_rv        boolean;
    l_ind       integer;
    l_idx       integer;
    l_info      tp_zip_info;
    l_name      raw(32767);
    l_utf8_name raw(32767);
  begin
    if p_name is null and p_idx is null
    then
      return false;
    end if;
    get_zip_info( p_zip, l_info );
    if nvl( l_info.cnt, 0 ) < 1
    then -- no (zip) file or empty zip file
      return false;
    end if;
    --
    if p_name is not null
    then
      l_name := char2raw( p_name, p_encoding );
      l_utf8_name := char2raw( p_name, 'AL32UTF8' );
    end if;
    --
    l_rv := false;
    l_ind := l_info.idx_cd;
    l_idx := 1;
    loop
      exit when not parse_central_file_header( p_zip, l_ind, p_cfh );
      if l_idx = p_idx
         or p_cfh.name1 = case when p_cfh.utf8 then l_utf8_name else l_name end
      then
        l_rv := true;
        exit;
      end if;
      l_ind := l_ind + p_cfh.len;
      l_idx := l_idx + 1;
    end loop;
    --
    p_cfh.idx := l_idx;
    p_cfh.encoding := get_encoding( p_encoding );
    return l_rv;
  end;
  --
  function parse_file( p_zipped_blob blob, p_fh in out tp_cfh )
  return blob
  is
    l_rv blob;
    l_buf raw(3999);
    l_compression_method varchar2(4);
    l_n integer;
    l_m integer;
    l_crc raw(4);
  begin
    if p_fh.original_len is null
    then
      raise_application_error( -20006, 'File not found' );
    end if;
    if nvl( p_fh.original_len, 0 ) = 0
    then
      return empty_blob();
    end if;
    l_buf := dbms_lob.substr( p_zipped_blob, 30, p_fh.offset + 1 );
    if utl_raw.substr( l_buf, 1, 4 ) != c_LOCAL_FILE_HEADER
    then
      raise_application_error( -20007, 'Error parsing the zipfile' );
    end if;
    l_compression_method := utl_raw.substr( l_buf, 9, 2 );
    l_n := little_endian( l_buf, 27, 2 );
    l_m := little_endian( l_buf, 29, 2 );
    if l_compression_method = '0800'
    then
      if p_fh.original_len < 32767 and p_fh.compressed_len < 32748
      then
        return utl_compress.lz_uncompress( utl_raw.concat
                 ( hextoraw( '1F8B0800000000000003' )
                 , dbms_lob.substr( p_zipped_blob, p_fh.compressed_len, p_fh.offset + 31 + l_n + l_m )
                 , p_fh.crc32
                 , utl_raw.substr( utl_raw.reverse( to_char( p_fh.original_len, 'fm0XXXXXXXXXXXXXXX' ) ), 1, 4 )
                 ) );
      end if;
      l_rv := hextoraw( '1F8B0800000000000003' ); -- gzip header
      dbms_lob.copy( l_rv
                   , p_zipped_blob
                   , p_fh.compressed_len
                   , 11
                   , p_fh.offset + 31 + l_n + l_m
                   );
      dbms_lob.append( l_rv
                     , utl_raw.concat( p_fh.crc32
                                     , utl_raw.substr( utl_raw.reverse( to_char( p_fh.original_len, 'fm0XXXXXXXXXXXXXXX' ) ), 1, 4 )
                                     )
                     );
      return utl_compress.lz_uncompress( l_rv );
    elsif l_compression_method = '0000'
    then
      if p_fh.original_len < 32767 and p_fh.compressed_len < 32767
      then
        return dbms_lob.substr( p_zipped_blob
                              , p_fh.compressed_len
                              , p_fh.offset + 31 + l_n + l_m
                              );
      end if;
      dbms_lob.createtemporary( l_rv, true, c_lob_duration );
      dbms_lob.copy( l_rv
                   , p_zipped_blob
                   , p_fh.compressed_len
                   , 1
                   , p_fh.offset + 31 + l_n + l_m
                   );
      return l_rv;
    end if;
    raise_application_error( -20008, 'Unhandled compression method ' || l_compression_method );
  end parse_file;
  --
  function get_count( p_zipped_blob blob )
  return integer
  is
    l_info tp_zip_info;
  begin
    get_zip_info( p_zipped_blob, l_info );
    return nvl( l_info.cnt, 0 );
  end;
  --
  function file2blob( p_dir varchar2, p_file_name varchar2 )
  return blob
  is
    file_lob bfile;
    file_blob blob;
    dest_offset integer := 1;
    src_offset  integer := 1;
  begin
    file_lob := bfilename( p_dir, p_file_name );
    dbms_lob.open( file_lob, dbms_lob.file_readonly );
    dbms_lob.createtemporary( file_blob, true, c_lob_duration );
    dbms_lob.loadblobfromfile( file_blob, file_lob, dbms_lob.lobmaxsize, dest_offset, src_offset );
    dbms_lob.close( file_lob );
    return file_blob;
  exception
    when others then
      if dbms_lob.isopen( file_lob ) = 1
      then
        dbms_lob.close( file_lob );
      end if;
      if dbms_lob.istemporary( file_blob ) = 1
      then
        dbms_lob.freetemporary( file_blob );
      end if;
      raise;
  end file2blob;
  --
  function get_sheet_names( p_xlsx blob )
  return sheet_names
  is
    l_cfh      tp_cfh;
    l_workbook blob;
    l_rv       sheet_names;
  begin
    if not get_central_file_header( p_xlsx, 'xl\workbook.xml', null, null, l_cfh )
    then
      for i in 1 .. get_count( p_xlsx )
      loop
      exit when not get_central_file_header( p_xlsx, null, i, null, l_cfh )
             or lower( utl_raw.cast_to_varchar2( l_cfh.name1 ) ) like '%workbook.xml';
      end loop;
    end if;
    if l_cfh.original_len is null
    then
      l_rv := sheet_names();
    else
      l_workbook := parse_file( p_xlsx, l_cfh );
      select xt1.name
      bulk collect into l_rv
      from xmltable( xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                                  , 'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x" )
                   , '( /workbook/sheets/sheet, /x:workbook/x:sheets/x:sheet )'
                     passing xmltype( xmldata => l_workbook, csid => nls_charset_id( 'AL32UTF8' ) )
                     columns
                       name varchar2(4000) path '@name'
                   ) xt1;
    end if;
    return l_rv;
  end get_sheet_names;
  --
  function read
    ( p_xlsx           blob
    , p_sheets         varchar2 := null
    , p_cell           varchar2 := null
    , p_include_clobs  varchar2 := null
    , p_add_empty_cols varchar2 := null
    )
  return tp_all_cells pipelined
  is
    l_nr             number;
    l_cnt            pls_integer;
    l_cfh            tp_cfh;
    l_name           varchar2(32767);
    l_workbook       blob;
    l_workbook_rels  blob;
    l_shared_strings blob;
    l_file           blob;
    l_csid_utf8      integer := nls_charset_id( 'AL32UTF8' );
    type tp_strings     is table of varchar2(32767);
    type tp_string_lens is table of varchar2(32767);
    type tp_boolean_tab is table of boolean index by pls_integer;
    l_strings      tp_strings;
    l_string_lens  tp_string_lens;
    l_quote_prefix tp_boolean_tab;
    l_date_styles  tp_boolean_tab;
    l_time_styles  tp_boolean_tab;
    l_one_cell     tp_one_cell;
    l_empty_cols    boolean;
    l_prev_col_nr   number(10);
    l_prev_row_nr   number(10);
    l_null_cell     tp_one_cell;
  begin
    l_cnt := get_count( p_xlsx );
    for i in 1 .. l_cnt
    loop
      exit when not get_central_file_header( p_xlsx, null, i, null, l_cfh );
      l_name := lower( utl_raw.cast_to_varchar2( l_cfh.name1 ) );
      if    l_name like '%workbook.xml'      then
        l_workbook      := parse_file( p_xlsx, l_cfh );
      elsif l_name like '%workbook.xml.rels' then
        l_workbook_rels := parse_file( p_xlsx, l_cfh );
      elsif l_name like '%sharedstrings.xml' then
        l_shared_strings := parse_file( p_xlsx, l_cfh );
        select xt1.txt
             , xt1.len
        bulk collect into l_strings, l_string_lens
        from xmltable( xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                                    , 'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x" )
                     , '( /sst/si, /x:sst/x:si )'
                       passing xmltype( xmldata => l_shared_strings, csid => l_csid_utf8 )
                       columns txt varchar2(4000 char) path 'substring( string-join(.//*:t/text(), "" ), 1, 3900 )'
                             , len integer             path 'string-length( string-join(.//*:t/text(), "" ) )'
                     ) xt1;
      elsif l_name like '%styles.xml' then
        l_file := parse_file( p_xlsx, l_cfh );
        for r_n in ( select xt2.seq - 1 seq
                          , xt2.id
                          , xt2.quoteprefix
                          , lower( xt3.format ) format
                     from xmltable( xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                                                 , 'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x" )
                                  , '( /styleSheet, /x:styleSheet )'
                                    passing xmltype( xmldata => l_file, csid => nls_charset_id( 'AL32UTF8' ) )
                                      columns cellxfs xmltype path '( cellXfs, x:cellXfs )'
                                            , numfmts xmltype path '( numFmts, x:numFmts )'
                                  ) xt1
                     cross join
                          xmltable( '/*:cellXfs/*:xf'
                                    passing xt1.cellxfs
                                      columns seq for ordinality
                                            , id          integer        path '@numFmtId'
                                            , quoteprefix varchar2(4000) path '@quotePrefix'
                                  ) xt2
                     left join
                          xmltable( '/*:numFmts/*:numFmt'
                                    passing xt1.numfmts
                                    columns id3    integer        path '@numFmtId'
                                          , format varchar2(4000) path '@formatCode'
                                  ) xt3
                     on xt3.id3 = xt2.id
                   )
        loop
          if    r_n.id between 14 and 17
             or instr( r_n.format, 'd' ) > 0
             or instr( r_n.format, 'y' ) > 0
          then
            l_date_styles( r_n.seq ) := null;
          elsif r_n.id between 18 and 22
             or r_n.id between 45 and 47
             or instr( r_n.format, 'h' ) > 0
             or instr( r_n.format, 'm' ) > 0
          then
            l_time_styles( r_n.seq ) := null;
          elsif r_n.quoteprefix in ( '1', 'true' )
          then
            l_quote_prefix( r_n.seq ) := null;
          end if;
        end loop;
        dbms_lob.freetemporary( l_file );
      end if;
    end loop;
    if l_workbook is null or l_workbook_rels is null
    then
      raise no_data_needed;
    end if;
    --
    if upper( substr( p_add_empty_cols, 1, 1 ) ) in ( 'Y', 'T', '1' )
    then
      l_empty_cols := true;
      l_null_cell.cell_type := 'S';
    end if;
    for r_x in ( select xt1.d1904
                      , xt2.seq
                      , xt2.name
                      , xt3.target
                 from xmltable( xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                                                     , 'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x" )
                              , '( /workbook, /x:workbook )'
                                passing xmltype( xmldata => l_workbook, csid => nls_charset_id( 'AL32UTF8' ) )
                                columns d1904  varchar2( 4000 ) path '*:workbookPr/@date1904'
                                      , sheets xmltype path '*:sheets'
                              ) xt1
                 cross join
                      xmltable( '*:sheets/*:sheet'
                                passing xt1.sheets
                                columns seq for ordinality
                                      , name    varchar2( 4000 ) path '@name'
                                      , sheetid varchar2( 4000 ) path '@sheetId'
                                      , rid     varchar2( 4000 ) path '@*:id[ namespace-uri(.) =
                                                                         ( "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                                                         , "http://purl.oclc.org/ooxml/officeDocument/relationships"
                                                                         ) ]'
                                      , state   varchar2( 4000 ) path '@state' ) xt2
                 join
                      xmltable( xmlnamespaces( default 'http://schemas.openxmlformats.org/package/2006/relationships' )
                              , '/Relationships/Relationship'
                                passing xmltype( xmldata => l_workbook_rels, csid => nls_charset_id( 'AL32UTF8' ) )
                                columns type    varchar2( 4000 ) path '@Type'
                                      , target  varchar2( 4000 ) path '@Target'
                                      , id      varchar2( 4000 ) path '@Id' ) xt3
                 on xt3.id = xt2.rid
                 order by xt2.sheetid
               )
    loop
      if ( p_sheets is null
         or instr( ':' || p_sheets || ':', ':' || r_x.seq || ':' ) > 0
         or instr( ':' || p_sheets || ':', ':' || r_x.name || ':' ) > 0
         )
      then
        for i in 1 .. l_cnt
        loop
          exit when not get_central_file_header( p_xlsx, null, i, null, l_cfh );
          if utl_raw.cast_to_varchar2( l_cfh.name1 ) like '%' || r_x.target
          then
            l_file := parse_file( p_xlsx, l_cfh );
            l_one_cell.sheet_nr   := r_x.seq;
            l_one_cell.sheet_name := r_x.name;
            l_one_cell.row_nr     := 0;
            l_prev_col_nr := 0;
            l_prev_row_nr := 0;
            l_null_cell.sheet_nr   := r_x.seq;
            l_null_cell.sheet_name := r_x.name;
            for r_c in ( select *
                         from xmltable( xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                                                     , 'http://purl.oclc.org/ooxml/spreadsheetml/main' as "x" )
                                      , '( /worksheet/sheetData/row/c, /x:worksheet/x:sheetData/x:row/x:c )'
                                        passing xmltype( xmldata => l_file, csid => l_csid_utf8 )
                                        columns v varchar2(4000) path '*:v'
                                              , f varchar2(4000) path '*:f'
                                              , t varchar2(4000) path '@t'
                                              , r varchar2(32)   path '@r'
                                              , s integer        path '@s'
                                              , rw integer       path './../@r'
                                              , txt varchar2(4000 char) path 'substring( string-join(.//*:t/text(), "" ), 1, 3900 )'
                                              , len integer             path 'string-length( string-join(.//*:t/text(), "" ) )'
                                      )
                       )
            loop
              if p_cell != r_c.r
              then
                continue;
              end if;
              if r_c.r is null
              then
                if l_one_cell.row_nr = r_c.rw
                then
                  l_one_cell.col_nr := l_one_cell.col_nr + 1;
                else
                  l_one_cell.col_nr := 1;
                end if;
              else
                l_one_cell.col_nr := col_alfan( r_c.r );
              end if;
              l_one_cell.row_nr     := coalesce( r_c.rw, l_one_cell.row_nr + 1 );
              l_one_cell.cell       := r_c.r;
              l_one_cell.formula    := r_c.f;
              l_one_cell.string_val := null;
              l_one_cell.number_val := null;
              l_one_cell.date_val   := null;
              l_one_cell.clob_val   := null;
              l_one_cell.string_len := null;
              if l_empty_cols
                 and (   l_one_cell.col_nr > l_prev_col_nr + 1
                     or (   l_one_cell.col_nr > 1
                        and l_one_cell.row_nr > l_prev_row_nr
                        )
                     )
              then
                l_null_cell.row_nr := l_one_cell.row_nr;
                for n in case
                           when l_one_cell.col_nr > l_prev_col_nr + 1
                             then l_prev_col_nr + 1
                             else 1
                         end .. l_one_cell.col_nr - 1
                loop
                  l_null_cell.col_nr := n;
                  l_null_cell.cell := alfan_col( n ) || l_null_cell.row_nr;
                  pipe row( l_null_cell );
                end loop;
              end if;
              if r_c.t = 's'
              then
                l_one_cell.cell_type := 'S';
                if r_c.v is not null
                then
                  l_one_cell.string_val := l_strings( to_number( r_c.v ) + 1 );
                  l_one_cell.string_len := l_string_lens( to_number( r_c.v ) + 1 );
                  if l_one_cell.string_len > 3900 and substr( p_include_clobs, 1, 1 ) in ( 'Y', 'y', '1' )
                  then
                    select xt1.txt
                    into l_one_cell.clob_val
                    from xmltable( '/*:sst/*:si[$i]'
                                   passing xmltype( xmldata => l_shared_strings, csid => l_csid_utf8 )
                                         , to_number( r_c.v ) + 1 as "i"
                                   columns txt clob path 'string-join(.//*:t/text(), "" )'
                                 ) xt1;
                  end if;
                  if l_quote_prefix.exists( r_c.s )
                  then
                    l_one_cell.string_val := '''' || l_one_cell.string_val;
                    if l_one_cell.clob_val is not null
                    then
                      l_one_cell.clob_val := '''' || l_one_cell.clob_val;
                    end if;
                  end if;
                end if;
              elsif r_c.t = 'n' or r_c.t is null
              then
                l_nr := to_number( r_c.v
                                 , case when instr( upper( r_c.v ), 'E' ) = 0
                                     then translate( r_c.v, '.012345678,-+', 'D999999999' )
                                     else translate( substr( r_c.v, 1, instr( upper( r_c.v ) , 'E' ) - 1 ), '.012345678,-+', 'D999999999' ) || 'EEEE'
                                   end
                                 , 'NLS_NUMERIC_CHARACTERS=.,'
                                 );
                if l_date_styles.exists( r_c.s )
                then
                  l_one_cell.cell_type := 'D';
                  if lower( r_x.d1904 ) in ( 'true', '1' )
                  then
                    l_one_cell.date_val := date '1904-01-01' + l_nr;
                  else
                    l_one_cell.date_val := date '1900-03-01' + ( l_nr - case when l_nr < 61 then 60 else 61 end );
                  end if;
                elsif l_time_styles.exists( r_c.s )
                then
                  l_one_cell.cell_type := 'S';
                  l_one_cell.string_val := to_char( numtodsinterval(  l_nr, 'day' ) );
                  l_one_cell.string_len := length( l_one_cell.string_val );
                else
                  l_one_cell.cell_type := 'N';
                  l_nr := round( l_nr, 14 - substr( to_char( l_nr, 'TME' ), -3 ) );
                  l_one_cell.number_val := l_nr;
                end if;
              elsif r_c.t = 'd'
              then
                l_one_cell.cell_type := 'D';
                l_one_cell.date_val := cast( to_timestamp_tz( r_c.v, 'yyyy-mm-dd"T"hh24:mi:ss.ffTZH:TZM' ) as date );
              elsif r_c.t = 'inlineStr'
              then
                l_one_cell.cell_type := 'S';
                l_one_cell.string_val := r_c.txt;
                l_one_cell.string_len := r_c.len;
                if     l_one_cell.string_len > 3900
                   and substr( p_include_clobs, 1, 1 ) in ( 'Y', 'y', '1' )
                   and r_c.r is not null
                then
                  select xt1.txt
                  into l_one_cell.clob_val
                  from xmltable( '/*:worksheet/*:sheetData/*:row/*:c[@r=$r]'
                                 passing xmltype( xmldata => l_file, csid => l_csid_utf8 )
                                       , r_c.r as "r"
                                 columns txt clob path 'string-join(.//*:t/text(), "" )'
                               ) xt1;
                end if;
              elsif r_c.t in ( 'str', 'e' )
              then
                l_one_cell.cell_type := 'S';
                l_one_cell.string_val := r_c.v;
                l_one_cell.string_len := length( l_one_cell.string_val );
              elsif r_c.t = 'b'
              then
                l_one_cell.cell_type := 'S';
                l_one_cell.string_val := case r_c.v
                                           when '1' then 'TRUE'
                                           when '0' then 'FALSE'
                                           else r_c.v
                                        end;
                l_one_cell.string_len := length( l_one_cell.string_val );
              end if;
              pipe row( l_one_cell );
              l_prev_col_nr := l_one_cell.col_nr;
              l_prev_row_nr := l_one_cell.row_nr;
            end loop;
            dbms_lob.freetemporary( l_file );
            exit;
          end if;
        end loop;
      end if;
    end loop;
    --
    dbms_lob.freetemporary( l_workbook );
    dbms_lob.freetemporary( l_workbook_rels );
    if l_strings is not null
    then
      l_strings.delete;
      l_string_lens.delete;
      dbms_lob.freetemporary( l_shared_strings );
    end if;
    l_date_styles.delete;
    l_time_styles.delete;
    l_quote_prefix.delete;
    raise no_data_needed;
  end read;
  --
  function get_version
  return varchar2
  is
  begin
    return c_version;
  end get_version;
  --
end as_xlsx;
/
