~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 5, 1, 5 );
  as_xlsx.cell( 3, 1, 3 );
  as_xlsx.cell( 2, 2, 45 );
  as_xlsx.cell( 3, 2, 'Anton Scheffer', p_alignment => as_xlsx.get_alignment( p_wraptext => true ) );
  as_xlsx.cell( 1, 4, sysdate, p_fontId => as_xlsx.get_font( 'Calibri', p_rgb => 'FFFF0000' ) );
  as_xlsx.cell( 2, 4, sysdate, p_numFmtId => as_xlsx.get_numFmt( 'dd/mm/yyyy h:mm' ) );
  as_xlsx.cell( 3, 4, sysdate, p_numFmtId => as_xlsx.get_numFmt( as_xlsx.orafmt2excel( 'dd/mon/yyyy' ) ) );
  as_xlsx.cell( 5, 5, 75, p_borderId => as_xlsx.get_border( 'double', 'double', 'double', 'double' ) );
  as_xlsx.cell( 2, 3, 33 );
  as_xlsx.hyperlink( 1, 6, 'http://www.amis.nl', 'Amis site' );
  as_xlsx.cell( 1, 7, 'Some merged cells', p_alignment => as_xlsx.get_alignment( p_horizontal => 'center' ) );
  as_xlsx.mergecells( 1, 7, 3, 7 );
  for i in 1 .. 5
  loop
    as_xlsx.comment( 3, i + 3, 'Row ' || (i+3), 'Anton' );
  end loop;
  as_xlsx.new_sheet;
  as_xlsx.set_row( 1, p_fillId => as_xlsx.get_fill( 'solid', 'FFFF0000' ) ) ;
  for i in 1 .. 5
  loop
    as_xlsx.cell( 1, i, i );
    as_xlsx.cell( 2, i, i * 3 );
    as_xlsx.cell( 3, i, 'x ' || i * 3 );
  end loop;
  as_xlsx.query2sheet( 'select rownum, x.*
, case when mod( rownum, 2 ) = 0 then rownum * 3 end demo
, case when mod( rownum, 2 ) = 1 then ''demo '' || rownum end demo2 from dual x connect by rownum <= 5' );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 6, 5 );
  as_xlsx.cell( 1, 7, 3 );
  as_xlsx.cell( 1, 8, 7 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 2, 6, 15, p_sheet => 2 );
  as_xlsx.cell( 2, 7, 13, p_sheet => 2 );
  as_xlsx.cell( 2, 8, 17, p_sheet => 2 );
  as_xlsx.list_validation( 6, 3, 1, 6, 1, 8, p_show_error => true, p_sheet => 1 );
  as_xlsx.defined_name( 2, 6, 2, 8, 'Anton', 2 );
  as_xlsx.list_validation
    ( 6, 1, 'Anton'
    , p_style => 'information'
    , p_title => 'valid values are'
    , p_prompt => '13, 15 and 17'
    , p_show_error => true
    , p_error_title => 'Are you sure?'
    , p_error_txt => 'Valid values are: 13, 15 and 17'
    , p_sheet => 1 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 6, 5 );
  as_xlsx.cell( 1, 7, 3 );
  as_xlsx.cell( 1, 8, 7 );
  as_xlsx.set_autofilter( 1,1, p_row_start => 5, p_row_end => 8 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 2, 6, 5 );
  as_xlsx.cell( 2, 7, 3 );
  as_xlsx.cell( 2, 8, 7 );
  as_xlsx.set_autofilter( 2,2, p_row_start => 5, p_row_end => 8 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.setUseXf( false );
  for c in 1 .. 10
  loop
    as_xlsx.cell( c, 1, 'COL' || c );
    as_xlsx.cell( c, 2, 'val' || c );
    as_xlsx.cell( c, 3, c );
  end loop;
  as_xlsx.freeze_rows( 1 );
  as_xlsx.new_sheet;
  for r in 1 .. 10
  loop
    as_xlsx.cell( 1, r, 'ROW' || r );
    as_xlsx.cell( 2, r, 'val' || r );
    as_xlsx.cell( 3, r, r );
  end loop;
  as_xlsx.freeze_cols( 3 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 3, 3, 'Start freeze' );
  as_xlsx.freeze_pane( 3,3 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
    select *
    from table( as_xlsx.read( as_xlsx.file2blob( 'MY_DIR', 'test.xlsx' ), '1' ) )
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.add_image( 1, 1, as_barcode.barcode( 'https://github.com/antonscheffer/as_xlsx', 'QR' ) );
  as_xlsx.cell( 1, 8, 'now with png images' );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
