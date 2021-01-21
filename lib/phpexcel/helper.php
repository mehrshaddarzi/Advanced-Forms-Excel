<?php

function createColumnsArray( $end_column, $first_letters = '' ) {
	$columns = array();
	$length  = strlen( $end_column );
	$letters = range( 'A', 'Z' );

	// Iterate over 26 letters.
	foreach ( $letters as $letter ) {

		// Paste the $first_letters before the next.
		$column = $first_letters . $letter;

		// Add the column to the final array.
		$columns[] = $column;

		// If it was the end column that was added, return the columns.
		if ( $column == $end_column ) {
			return $columns;
		}
	}

	// Add the column children.
	foreach ( $columns as $column ) {
		// Don't itterate if the $end_column was already set in a previous itteration.
		// Stop iterating if you've reached the maximum character length.
		if ( ! in_array( $end_column, $columns ) && strlen( $column ) < $length ) {
			$new_columns = createColumnsArray( $end_column, $column );
			// Merge the new columns which were created with the final columns array.
			$columns = array_merge( $columns, $new_columns );
		}
	}


	//start with ket=>1 in php
	$export_array = $columns;
	$export_array = array_filter( array_merge( array( 0 ), $export_array ) );


	return $export_array;
}

/* PHPExcel export */
function phpexcel( $sheetname, $field, $data, $file_pishvand = false, $auto_row = false ) {

	/*Example
	$f = array(
		array("name" => "esm", "size" => "", "link" => "yes"),
		array("name" => "email", "size" => "auto", "link" => "no"),
		array("name" => "سال", "size" => "auto", "link" => "no"),
	);
	$data = array(
	array("View profile|||http://www.irwebdesign.ir","mehrshadhjg198@gmail.com","28"),
	array("View profile|||http://www.mehdi.net","mehrshad198@gmail.com","45"),
	);

	echo phpexcel("me",$f,$data);
	*/

	// Create new PHPExcel object
	$objPHPExcel = new PHPExcel();

	// Set document properties
	$objPHPExcel->getProperties()->setCreator( "Maarten Balliauw" )
	            ->setLastModifiedBy( "Maarten Balliauw" )
	            ->setTitle( "Office 2007 XLSX Test Document" )
	            ->setSubject( "Office 2007 XLSX Test Document" )
	            ->setDescription( "Test document for Office 2007 XLSX, generated using PHP classes." )
	            ->setKeywords( "office 2007 openxml php" )
	            ->setCategory( "Test result file" );

	// Rename worksheet
	$objPHPExcel->getActiveSheet()->setTitle( $sheetname );
	$objPHPExcel->getActiveSheet()->setRightToLeft( true );
	$objPHPExcel->getDefaultStyle()->getFont()->setName( 'tahoma' )->setSize( 11 );
	$objPHPExcel->getActiveSheet()
	            ->getStyle( 'A1:CF1' )
	            ->applyFromArray(
		            array(
			            'fill' => array(
				            'type'  => PHPExcel_Style_Fill::FILL_SOLID,
				            'color' => array( 'rgb' => '18A689' )
			            ),
		            )
	            );
	$objPHPExcel->getActiveSheet()->getStyle( 'A1:CF1' )->applyFromArray( array( 'font' => array( 'color' => array( 'rgb' => 'ffffff' ) ) ) );
	$objPHPExcel->setActiveSheetIndex( 0 );

	//Array
	$arr  = createColumnsArray( 'CC' );
	$link = array();
	$head = array();

	if ( $auto_row ) {
		$head[] = $auto_row;
		$x      = 1;
	} else {
		$x = 0;
	}

	foreach ( $field as $header ) {
		$head[] = $header['name'];
		if ( $header['link'] == "yes" ) {
			$link[] = $x;
		}
		$x ++;
	}
	$objPHPExcel->getActiveSheet()->fromArray( $head, null, 'A1' );

	//add auto number in data
	if ( $auto_row ) {
		for ( $m = 0; $m < count( $data ); $m ++ ) {
			array_unshift( $data[ $m ], ( $m + 1 ) );
		}
	}

	//replace link data
	if ( count( $link > 0 ) ) {
		$list_link = array();
		for ( $z = 0; $z < count( $data ); $z ++ ) {
			foreach ( $link as $p ) {
				$new_array        = explode( "|||", $data[ $z ][ $p ] ); //[0] view [1] link
				$list_link[ $z ]  = $new_array[1];
				$data[ $z ][ $p ] = $new_array[0];
			}
		}
	}

	$objPHPExcel->getActiveSheet()->fromArray( $data, null, 'A2' );

	//set link for data
	if ( count( $link > 0 ) ) {
		for ( $q = 1; $q <= count( $data ); $q ++ ) {
			foreach ( $link as $p ) {

				//if auto number
				$objPHPExcel->getActiveSheet()->getCell( $arr[ $p + 1 ] . ( $q + 1 ) )->getHyperlink()->setUrl( $list_link[ $q - 1 ] );
				$objPHPExcel->getActiveSheet()->getStyle( $arr[ $p + 1 ] . ( $q + 1 ) )->applyFromArray( array( 'font' => array( 'color' => array( 'rgb' => '0000ff' ) ) ) );

			}
		}
	}

	//setu auto size
	if ( $auto_row ) {
		$objPHPExcel->getActiveSheet()->getColumnDimension( $arr[1] )->setAutoSize( true );
		$t = 2;
	} else {
		$t = 1;
	}
	foreach ( $field as $header ) {
		if ( $header['size'] == "auto" ) {
			$objPHPExcel->getActiveSheet()->getColumnDimension( $arr[ $t ] )->setAutoSize( true );
		}
		$t ++;
	}
	$objWriter = PHPExcel_IOFactory::createWriter( $objPHPExcel, 'Excel5' );

	//save file
	$upload_dir = wp_upload_dir(); // Grab uploads folder array
	$dir        = trailingslashit( $upload_dir['basedir'] ) . 'excel/'; // Set storage directory path
	if ( ! file_exists( $dir ) ) {
		wp_mkdir_p( $dir );
	}

	//Remove All file From this directory
	foreach ( glob( $dir . '*.*' ) as $v ) {
		if ( is_file( $v ) ) {
			if ( time() - filemtime( $v ) >= 60 * 60 * 24 * 2 ) { // 2 days
				@unlink( $v );
			}
		}
	}

	if ( $file_pishvand ) {
		$pish = $file_pishvand;
	} else {
		$pish = "excel";
	}
	$file_name = $pish . "-" . date( 'Y-m-d-His' ) . ".xls";

	$objWriter->save( $dir . $file_name );

	// Return file Link
	return $upload_dir['baseurl'] . '/excel/' . $file_name;
}