# Import the xlrd module
import xlrd
import pymysql
file = r'path_to_file\Excell_all.xlsx'
book = xlrd.open_workbook(file)
sheet = book.sheet_by_name("out_")

connection = pymysql.connect(host='localhost',
                     user='root',
                     password='Pass_to_MySQLDB',
                     database='new_data')
query = """INSERT INTO alldata (sku, store_view_code, attribute_set_code, product_type, categories, product_websites, name, description,
                             short_description, weight, product_online, tax_class_name, visibility, price, special_price,
                             special_price_from_date, special_price_to_date, url_key, meta_title, meta_keywords, meta_description, base_image,
                             base_image_label, small_image, small_image_label, thumbnail_image, thumbnail_image_label, swatch_image,
                             swatch_image_label, created_at, updated_at, new_from_date, new_to_date, display_product_options_in, map_price,
                             msrp_price, map_enabled, gift_message_available, custom_design, custom_design_from, custom_design_to,
                             custom_layout_update, page_layout, product_options_container, msrp_display_actual_price_type, country_of_manufacture,
                             additional_attributes, qty, out_of_stock_qty, use_config_min_qty, is_qty_decimal, allow_backorders,
                             use_config_backorders, min_cart_qty, use_config_min_sale_qty, max_cart_qty, use_config_max_sale_qty,
                             is_in_stock, notify_on_stock_below, use_config_notify_stock_qty, manage_stock, use_config_manage_stock,
                             use_config_qty_increments, qty_increments, use_config_enable_qty_inc, enable_qty_increments,
                             is_decimal_divided, website_id, related_skus, related_position, crosssell_skus, crosssell_position,
                             upsell_skus, upsell_position, additional_images, additional_image_labels, hide_from_product_page,
                             custom_options, bundle_price_type, bundle_sku_type, bundle_price_view, bundle_weight_type, bundle_values,
                             bundle_shipment_type, associated_skus, configurable_variations, configurable_variation_labels) 
            VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                   %s, %s, %s, %s, %s, %s, %s)"""

cursor = connection.cursor()
for r in range(1, sheet.nrows):
    sku		                   = sheet.cell(r, 0).value
    store_view_code	           = sheet.cell(r, 1).value
    attribute_set_code         = sheet.cell(r, 2).value
    product_type               = sheet.cell(r, 3).value
    categories                 = sheet.cell(r, 4).value
    product_websites           = sheet.cell(r, 5).value
    name                       = sheet.cell(r, 6).value
    description                = sheet.cell(r, 7).value
    short_description          = sheet.cell(r, 8).value
    weight                     = sheet.cell(r, 9).value
    product_online             = sheet.cell(r, 10).value
    tax_class_name             = sheet.cell(r, 11).value
    visibility                 = sheet.cell(r, 12).value
    price                      = sheet.cell(r, 13).value
    special_price              = sheet.cell(r, 14).value
    special_price_from_date    = sheet.cell(r, 15).value
    special_price_to_date      = sheet.cell(r, 16).value
    url_key                    = sheet.cell(r, 17).value
    meta_title                 = sheet.cell(r, 18).value
    meta_keywords              = sheet.cell(r, 19).value
    meta_description           = sheet.cell(r, 20).value
    base_image                 = sheet.cell(r, 21).value
    base_image_label           = sheet.cell(r, 22).value
    small_image                = sheet.cell(r, 23).value
    small_image_label          = sheet.cell(r, 24).value
    thumbnail_image            = sheet.cell(r, 25).value
    thumbnail_image_label      = sheet.cell(r, 26).value
    swatch_image               = sheet.cell(r, 27).value
    swatch_image_label         = sheet.cell(r, 28).value
    created_at                 = sheet.cell(r, 29).value
    updated_at                 = sheet.cell(r, 30).value
    new_from_date              = sheet.cell(r, 31).value
    new_to_date                = sheet.cell(r, 32).value
    display_product_options_in = sheet.cell(r, 33).value
    map_price                  = sheet.cell(r, 34).value
    msrp_price                 = sheet.cell(r, 35).value
    map_enabled                = sheet.cell(r, 36).value
    gift_message_available     = sheet.cell(r, 37).value
    custom_design              = sheet.cell(r, 38).value
    custom_design_from         = sheet.cell(r, 39).value
    custom_design_to           = sheet.cell(r, 40).value
    custom_layout_update       = sheet.cell(r, 41).value
    page_layout                = sheet.cell(r, 42).value
    product_options_container  = sheet.cell(r, 43).value
    msrp_display_actual_price_type = sheet.cell(r, 44).value
    country_of_manufacture         = sheet.cell(r, 45).value
    additional_attributes          = sheet.cell(r, 46).value
    qty                            = sheet.cell(r, 47).value
    out_of_stock_qty               = sheet.cell(r, 48).value
    use_config_min_qty             = sheet.cell(r, 49).value
    is_qty_decimal                 = sheet.cell(r, 50).value
    allow_backorders               = sheet.cell(r, 51).value
    use_config_backorders          = sheet.cell(r, 52).value
    min_cart_qty                   = sheet.cell(r, 53).value
    use_config_min_sale_qty        = sheet.cell(r, 54).value
    max_cart_qty                   = sheet.cell(r, 55).value
    use_config_max_sale_qty        = sheet.cell(r, 56).value
    is_in_stock                    = sheet.cell(r, 57).value
    notify_on_stock_below          = sheet.cell(r, 58).value
    use_config_notify_stock_qty    = sheet.cell(r, 59).value
    manage_stock                   = sheet.cell(r, 60).value
    use_config_manage_stock        = sheet.cell(r, 61).value
    use_config_qty_increments      = sheet.cell(r, 62).value
    qty_increments                 = sheet.cell(r, 63).value
    use_config_enable_qty_inc      = sheet.cell(r, 64).value
    enable_qty_increments          = sheet.cell(r, 65).value
    is_decimal_divided             = sheet.cell(r, 66).value
    website_id                     = sheet.cell(r, 67).value
    related_skus                   = sheet.cell(r, 68).value
    related_position               = sheet.cell(r, 69).value
    crosssell_skus                 = sheet.cell(r, 70).value
    crosssell_position             = sheet.cell(r, 71).value
    upsell_skus                    = sheet.cell(r, 72).value
    upsell_position                = sheet.cell(r, 73).value
    additional_images              = sheet.cell(r, 74).value
    additional_image_labels        = sheet.cell(r, 75).value
    hide_from_product_page         = sheet.cell(r, 76).value
    custom_options                 = sheet.cell(r, 77).value
    bundle_price_type              = sheet.cell(r, 78).value
    bundle_sku_type                = sheet.cell(r, 79).value
    bundle_price_view              = sheet.cell(r, 80).value
    bundle_weight_type             = sheet.cell(r, 81).value
    bundle_values                  = sheet.cell(r, 82).value
    bundle_shipment_type           = sheet.cell(r, 83).value
    associated_skus                = sheet.cell(r, 84).value
    configurable_variations        = sheet.cell(r, 85).value
    configurable_variation_labels  = sheet.cell(r, 86).value

    values = (sku, store_view_code, attribute_set_code, product_type, categories, product_websites, name, description,
                             short_description, weight, product_online, tax_class_name, visibility, price, special_price,
                             special_price_from_date, special_price_to_date, url_key, meta_title, meta_keywords, meta_description, base_image,
                             base_image_label, small_image, small_image_label, thumbnail_image, thumbnail_image_label, swatch_image,
                             swatch_image_label, created_at, updated_at, new_from_date, new_to_date, display_product_options_in, map_price,
                             msrp_price, map_enabled, gift_message_available, custom_design, custom_design_from, custom_design_to,
                             custom_layout_update, page_layout, product_options_container, msrp_display_actual_price_type, country_of_manufacture,
                             additional_attributes, qty, out_of_stock_qty, use_config_min_qty, is_qty_decimal, allow_backorders,
                             use_config_backorders, min_cart_qty, use_config_min_sale_qty, max_cart_qty, use_config_max_sale_qty,
                             is_in_stock, notify_on_stock_below, use_config_notify_stock_qty, manage_stock, use_config_manage_stock,
                             use_config_qty_increments, qty_increments, use_config_enable_qty_inc, enable_qty_increments,
                             is_decimal_divided, website_id, related_skus, related_position, crosssell_skus, crosssell_position,
                             upsell_skus, upsell_position, additional_images, additional_image_labels, hide_from_product_page,
                             custom_options, bundle_price_type, bundle_sku_type, bundle_price_view, bundle_weight_type, bundle_values,
                             bundle_shipment_type, associated_skus, configurable_variations, configurable_variation_labels)
    cursor.execute(query, values)


connection.commit()
cursor.close()
connection.close()

