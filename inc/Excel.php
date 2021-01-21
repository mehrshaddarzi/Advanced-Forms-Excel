<?php

namespace Advanced_Forms_Excel;

use Advanced_Forms_Excel\core\Utility;

/**
 * Excel Export Process
 */
class Excel_Export
{

    public function __construct()
    {
        add_action('admin_init', array($this, 'process_excel_export'), 0);
        add_action('admin_notices', array($this, 'my_test_plugin_admin_notice'));
    }

    public function my_test_plugin_admin_notice()
    {
        global $pagenow;
        if (($pagenow == "edit.php" || $pagenow == "users.php") and isset($_GET['alert_excel_export']) and isset($_GET['file']) and isset($_GET['rows'])) {
            if ($_GET['rows'] > 0) {
                $text = ' خروجی اکسل با ' . number_format_i18n($_GET['rows']) . ' ردیف با موفقیت ایجاد شد ';
                $text .= '&nbsp;<a href="' . trim($_GET['file']) . '" style="color: blue; text-decoration: none;">دانلود فایل اکسل</a>';
                Utility::admin_notice($text, 'success');
            } else {
                Utility::admin_notice("هیچ رکورد یافت نشد", 'error');
            }
        }
    }

    public function process_excel_export()
    {
        if (isset($_POST['excel_export_nonce'])
            and wp_verify_nonce($_POST['excel_export_nonce'], 'excel_export_nonce')
            and isset($_POST['post_type'])
            and isset($_POST['screen_id'])
        ) {

            // post Type
            $post_type = trim($_POST['post_type']);
            $screen_id = trim($_POST['screen_id']);
            if (empty($post_type)) {
                $post_type = $screen_id; //users
            }

            // Get Directory
            $upload_dir = wp_upload_dir(null, false);

            // Get Default Path
            $defaultPath = rtrim($upload_dir['basedir'], "/") . '/' . 'excel/';
            $default_link = rtrim($upload_dir['baseurl'], "/") . '/' . 'excel/';
            if (!file_exists($defaultPath)) {
                @mkdir($defaultPath, 0777, true);
            }

            // Include PHPExcel
            require_once \Advanced_Forms_Excel::$plugin_path . '/lib/phpexcel/core/PHPExcel.php';

            // Load PHP Excel Helper
            require_once \Advanced_Forms_Excel::$plugin_path . '/lib/phpexcel/helper.php';

            // List Field Excel
            // https://gist.github.com/mehrshaddarzi/7a32097bdf9cc686b7fa5bff365e126e
            $col = array(
                array("name" => "شناسه", "size" => "auto", "link" => "no")
            );
            $excel_col = apply_filters('excel_export_post_type_' . $post_type . '_col', $col);

            // Data Of Excel
            $data = array();
            $excel_data = apply_filters('excel_export_post_type_' . $post_type . '_data', $data);

            // Export
            $auto_excel_row = apply_filters('excel_export_post_type_' . $post_type . '_auto_row', false);
            $file_link = '';
            if ($GLOBALS['count_number_row_excel'] > 0) {
                $file_link = phpexcel('لیست', $excel_col, $excel_data, $file_prefix = 'excel', $auto_excel_row);
            }

            // Redirect to Show Notice
            if (empty($_POST['post_type'])) {
                $link = 'users.php?alert_excel_export=yes&file=' . $file_link . '&rows=' . $GLOBALS['count_number_row_excel'];
            } else {
                $link = 'edit.php?post_type=' . $_POST['post_type'] . '&alert_excel_export=yes&file=' . $file_link . '&rows=' . $GLOBALS['count_number_row_excel'];
            }
            wp_redirect(admin_url($link));
            exit;
        }
    }
}


new Excel_Export();