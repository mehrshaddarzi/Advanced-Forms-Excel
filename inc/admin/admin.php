<?php

namespace Advanced_Forms_Excel\admin;

use Advanced_Forms_Excel;

class Admin
{

    public static $af_form_post_type = 'af_form';
    public static $af_entry_post_type = 'af_entry';

    public function __construct()
    {

        /*add_action('admin_init', function () {
            if (isset($_GET['_test'])) {
                echo '<pre>';
                $form_key = 'form_5f4558083e515';
                $field_keys = self::getFormFields($form_key, true);
                $getEntryList = self::getEntryList($form_key);
                foreach ($getEntryList as $entry_id) {
                    echo '<pre>';
                    var_dump(self::getEntryDetailByID($entry_id, $field_keys));
                }
                exit;
            }
        });*/


        // Add Excel Export System
        add_action('admin_enqueue_scripts', array($this, 'admin_assets'));
        add_action('admin_footer', array($this, 'admin_assets_html_excel'));
        add_action('custom_excel_export_post_type_fields', array($this, 'custom_excel_export_fields'));
        add_filter('excel_export_post_type_' . self::$af_form_post_type . '_col', array($this, 'excel_col'));
        add_filter('excel_export_post_type_' . self::$af_form_post_type . '_data', array($this, 'excel_data'));

    }


    public function admin_assets()
    {
        global $pagenow;
        if ($pagenow == "edit.php" and isset($_GET['post_type']) and $_GET['post_type'] == self::$af_form_post_type) {
            add_thickbox();
        }
    }

    public function admin_assets_html_excel()
    {
        global $pagenow;
        if ($pagenow == "edit.php" and isset($_GET['post_type']) and $_GET['post_type'] == self::$af_form_post_type) {
            $form_list = self::getFormsList();
            if (count($form_list) > 0) {
                include \Advanced_Forms_Excel::$plugin_path . '/inc/excel-export.php';
            }
        }
    }

    public function excel_col($col)
    {
        $col = array(
            array("name" => "شناسه سیستم", "size" => "auto", "link" => "no"),
            array("name" => "تاریخ ثبت", "size" => "auto", "link" => "no"),
            array("name" => "ساعت ثبت", "size" => "auto", "link" => "no")
        );

        // Add ACf Field
        if (isset($_POST['form_key'])) {
            $form_key = $_POST['form_key'];
            $field_keys = self::getFormFields($form_key);
            foreach ($field_keys as $field) {
                $col[] = array("name" => $field['label'], "size" => "auto", "link" => "no");
            }
        }

        return $col;
    }

    public function excel_data($data)
    {
        global $wpdb;
        $arg = array(
            'post_type' => self::$af_entry_post_type,
            'post_status' => 'publish',
            'meta_query' => array(
                array(
                    'key' => 'form_key',
                    'value' => $_POST['form_key'],
                    'compare' => '='
                )
            ),
            'orderby' => 'ID',
            'order' => $_POST['order'],
            'offset' => $_POST['offset'] - 1,
        );

        // Check Number
        if ($_POST['posts_per_page'] != "ALL") {
            $arg['posts_per_page'] = $_POST['posts_per_page'];
        }

        // Get Form Field
        $form_fields_keys = self::getFormFields($_POST['form_key'], true);

        // Prepare Data
        $data = array();
        $query = Advanced_Forms_Excel\core\Utility::wp_query($arg, false);
        $GLOBALS['count_number_row_excel'] = count($query);
        foreach ($query as $post_ID) {
            $field = array();
            $post = get_post($post_ID);

            // ID
            $field[] = $post_ID;

            // Date
            $field[] = date_i18n('Y-m-d', strtotime($post->post_date));

            // Hour
            $field[] = date_i18n('H:i', strtotime($post->post_date));

            // Custom Field
            $entry_detail = self::getEntryDetailByID($post_ID, $form_fields_keys);
            foreach ($form_fields_keys as $f) {
                $field[] = $entry_detail[$f];
            }

            // Push To List
            $data[] = $field;
        }

        return $data;
    }

    public function custom_excel_export_fields($post_type)
    {
        if ($post_type == self::$af_form_post_type) {
            $form_list = self::getFormsList();
            ?>
            <tr valign="top">
                <td scope="row"><label for="tablecell">فرم</label></td>
                <td>
                    <select name="form_key">
                        <?php
                        foreach ($form_list as $array) {
                            $entry_list = self::getEntryList($array['key']);
                            if (count($entry_list) < 1) {
                                continue;
                            }
                            ?>
                            <option value="<?php echo $array['key']; ?>"><?php echo $array['title']; ?></option>
                            <?php
                        }
                        ?>
                    </select>
                </td>
            </tr>
            <?php
        }
    }

    public static function getFormFields($form_key, $only_key = false)
    {
        //$form = af_get_form( $post->ID );
        $field_groups = af_get_form_field_groups($form_key);
        $field_key = array();
        /**
         * array(4) {
         * [0]=>
         * array(3) {
         * ["label"]=>
         * string(6) "نام"
         * ["name"]=>
         * string(4) "name"
         * ["type"]=>
         * string(6) "متن"
         * }
         * [1]=>
         * array(3) {
         * ["label"]=>
         * string(21) "شماره همراه"
         * ["name"]=>
         * string(6) "mobile"
         * ["type"]=>
         * string(6) "متن"
         * }
         * [2]=>
         * array(3) {
         * ["label"]=>
         * string(10) "موضوع"
         * ["name"]=>
         * string(8) "category"
         * ["type"]=>
         * string(6) "متن"
         * }
         * [3]=>
         * array(3) {
         * ["label"]=>
         * string(15) "متن پیام"
         * ["name"]=>
         * string(7) "content"
         * ["type"]=>
         * string(38) "جعبه متن (متن چند خطی)"
         * }
         * }
         */

        $list = array();
        foreach ($field_groups as $field_group) {
            $fields = acf_get_fields($field_group);
            foreach ($fields as $field) {
                $field_key[] = $field['name'];
                $list[] = array(
                    'label' => $field['label'],
                    'name' => $field['name'],
                    'type' => acf_get_field_type_label($field['type'])
                );
            }
        }

        if ($only_key) {
            return $field_key;
        }

        return $list;
    }

    public static function getEntryList($form_key)
    {
        return Advanced_Forms_Excel\core\Utility::wp_query(array(
            'post_type' => self::$af_entry_post_type,
            'meta_query' => array(
                array(
                    'key' => 'form_key',
                    'value' => $form_key,
                    'compare' => '='
                )
            )
        ), false);
    }

    public static function getFormsList()
    {
        $list = array();
        $post_ids = Advanced_Forms_Excel\core\Utility::wp_query(array(
                'post_type' => self::$af_form_post_type
            )
        );
        foreach ($post_ids as $post_id => $title) {
            $list[] = array(
                'ID' => $post_id,
                'title' => $title,
                'key' => get_post_meta($post_id, 'form_key', true),
                'entry_number' => count(self::getEntryList(get_post_meta($post_id, 'form_key', true))),
                'meta' => get_post_meta($post_id)
            );
        }

        return $list;
    }

    public static function getEntryDetailByID($post_id, $meta_list = array())
    {
        $list = array();
        $post_metas = get_post_meta($post_id);
        $post_metas = array_combine(array_keys($post_metas), array_column($post_metas, '0'));
        $post = get_post($post_id);

        $list['ID'] = $post_id;
        $list['post_date'] = $post->post_date;
        foreach ($meta_list as $meta_key) {
            if (isset($post_metas[$meta_key])) {
                $list[$meta_key] = $post_metas[$meta_key];
            }
        }

        return $list;
    }

}