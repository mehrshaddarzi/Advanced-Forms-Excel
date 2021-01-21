<?php
global $post_type;
?>
<script>
    jQuery(document).ready(function ($) {
        // Add TickBox After Insert Post
        // @see https://codeontrack.com/wordpress-thickbox-use-modal-lightboxes-in-the-admin-area/
        // @see <a href='#TB_inline?height=400&amp;width=500&amp;inlineId=itunes_api' title='Select Bar Style' class='thickbox page-title-action'>iTunes API</a>
        $(`<a href='#TB_inline?height=600&amp;width=550&amp;inlineId=excel_export_tick_box' class='thickbox page-title-action'>خروجی اکسل</a>`).insertAfter("a.page-title-action");
    });
</script>
<div id="excel_export_tick_box" style="display:none;">
    <form action="" method="post">
        <table id="excel-export-table" class="form-table" dir="rtl" style="direction: rtl;text-align: right;">

            <!-- Order -->
            <tr valign="top">
                <td scope="row"><label for="tablecell">ترتیب</label></td>
                <td>
                    <select name="order">
                        <option value="ASC">صعودی</option>
                        <option value="DESC">نزولی</option>
                    </select>
                </td>
            </tr>

            <!-- Number -->
            <tr valign="top">
                <td scope="row"><label for="tablecell">تعداد خروجی</label></td>
                <td>
                    <select name="posts_per_page">
                        <option value="200">200</option>
                        <option value="500">500</option>
                        <option value="1000">1000</option>
                        <option value="5000">5000</option>
                        <option value="ALL">همه</option>
                    </select>
                </td>
            </tr>

            <!-- offset -->
            <tr valign="top">
                <td scope="row"><label for="tablecell">شروع از رکورد</label></td>
                <td>
                    <input type="number" name="offset" value="1" min="1" required/>
                </td>
            </tr>

            <!-- Custom Field -->
            <?php
            $screen = get_current_screen();
            do_action('custom_excel_export_post_type_fields', $post_type, $screen->id);
            ?>

            <!-- Post Type -->
            <input type="hidden" name="post_type" value="<?php echo $post_type; ?>">
            <input type="hidden" name="screen_id" value="<?php echo $screen->id; ?>">

            <!-- Nonce -->
            <?php wp_nonce_field('excel_export_nonce', 'excel_export_nonce'); ?>

            <!-- Submit Button -->
            <tr valign="top">
                <td scope="row"><?php submit_button('دریافت اکسل'); ?></td>
                <td></td>
            </tr>
        </table>
    </form>
</div>
<style>
    #excel-export-table td {
        margin-bottom: 5px !important;
        padding: 10px 10px !important;
    }
</style>