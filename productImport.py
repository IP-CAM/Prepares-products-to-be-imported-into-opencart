import xlrd
import datetime
import codecs

file_location = "C:/Users/sale/Documents/ChemFarmImports/productImportExcel/productImportExample2.xlsx"
#file_location = "C:/Users/brian.gao/Downloads/cFarm/productImport/productImportExample2.xlsx"

workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)

index = 0

product_id = sheet.col_values(index,1)
index += 1
product_option_id = sheet.col_values(index,1)
index += 1
name = sheet.col_values(index,1)
index += 1
description = sheet.col_values(index,1)
index += 1
meta_title = sheet.col_values(index,1)
index += 1
meta_description = sheet.col_values(index,1)
index += 1
meta_keyword = sheet.col_values(index,1)
index += 1
product_option_value_id = sheet.col_values(index,1)[0]
index += 1
option_value_id = sheet.col_values(index,1)
index += 1
quantities = sheet.col_values(index,1)
index += 1
units = sheet.col_values(index,1)
index += 1
prices = sheet.col_values(index,1)
index += 1
image = sheet.col_values(index,1)
index += 1
library = sheet.col_values(index,1)
index += 1
library_base_price = sheet.col_values(index,1)
index += 1
attribute_ids = sheet.col_values(index,1)
index += 1

texts = sheet.col_values(index,1)
index += 1

categories = sheet.col_values(index,1)
index += 1

SEO = sheet.col_values(index,1)

file_name = "productImport.sql"

#file = open("../productImportSQL/" + file_name,"w")


# delete
delete_oc_product = "DEDELETE FROM `oc_product` WHERE "
delete_oc_product_description = "DELETE FROM `oc_product_description` WHERE "
delete_oc_product_to_store = "DELETE FROM `oc_product_to_store` WHERE "
delete_oc_product_attribute = "DELETE FROM `oc_product_attribute` WHERE "
delete_oc_product_image = "DELETE FROM `oc_product_image` WHERE "
delete_oc_product_to_category = "DELETE FROM `oc_product_to_category` WHERE "
delete_oc_product_to_category2 = "DELETE FROM `oc_product_to_category2` WHERE "
delete_oc_product_to_layout = "DELETE FROM `oc_product_to_layout` WHERE "
delete_oc_url_alias = "DELETE FROM `oc_url_alias` WHERE "


delete_oc_product_option = "DELETE FROM `oc_product_option` WHERE "
delete_oc_product_option_value = "DELETE FROM `oc_product_option_value` WHERE "

# insert
oc_product = "INSERT INTO `oc_product` (`product_id`,`model`,`quantity`,`stock_status_id`,`image`,`shipping`,`date_available`,`weight_class_id`,`length_class_id`,`subtract`,`minimum`,`status`,`date_added`,`date_modified`,`library`,`library_base_price`) VALUES\n"
oc_product_description = "INSERT INTO `oc_product_description` (`product_id`,`language_id`,`name`,`description`,`meta_title`,`meta_description`,`meta_keyword`) VALUES\n"
oc_product_to_store = "INSERT INTO `oc_product_to_store` (`product_id`,`store_id`) VALUES\n"
oc_product_attribute = "INSERT INTO `oc_product_attribute` (`product_id`,`attribute_id`,`language_id`,`text`) VALUES\n"
##oc_product_discount = "INSERT INTO `oc_product_discount`"
##oc_product_special = "INSERT INTO `oc_product_special`"
oc_product_image = "INSERT INTO `oc_product_image` (`product_id`,`image`) VALUES\n"
#oc_product_to_download = "INSERT INTO `oc_product_to_download`"
oc_product_to_category = "INSERT INTO `oc_product_to_category` (`product_id`,`category_id`) VALUES\n"
oc_product_to_category2 = "INSERT INTO `oc_product_to_category2` (`product_id`,`category_id`) VALUES\n"
##oc_product_filter = "INSERT INTO `oc_product_filter`"
##oc_product_related = "INSERT INTO `oc_product_related`"
##oc_product_reward = "INSERT INTO `oc_product_reward`"
oc_product_to_layout = "INSERT INTO `oc_product_to_layout` (`product_id`) VALUES\n"
#oc_product_recurring = "INSERT INTO `oc_product_recurring`"
oc_url_alias ="INSERT INTO `oc_url_alias` (`query`,`keyword`) VALUES\n"

oc_product_option = "INSERT INTO `oc_product_option` (`product_id`,`option_id`) VALUES\n"
oc_product_option_value = "INSERT INTO `oc_product_option_value` (`product_option_value_id`,`product_option_id`,`option_id`,`product_id`,`option_value_id`,`table_unit`,`table_quantity`,`table_price`) VALUES\n"

# iterations
for i in range(len(product_id)-1):
    # delete
    delete_oc_product += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_description += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_to_store += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_attribute += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_image += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_to_category += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_to_category2 += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_to_layout += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_url_alias += "query='product_id=" + str(product_id[i]) + "' OR "

    delete_oc_product_option += "product_id='" + str(product_id[i]) + "' OR "
    delete_oc_product_option_value += "product_id='" + str(product_id[i]) + "' OR "
    
    # oc_product
    oc_product += "('" + str(product_id[i]) + "','Molecules',999,6,'catalog/" + str(image[i]) + "',1,'" + datetime.datetime.today().strftime('%Y-%m-%d') + "',1,1,1,1,1,NOW(),NOW(),'" + str(library[i]) + "','" + str(library_base_price[i]) + "'),\n"

    # oc_product_description
    oc_product_description += "('" + str(product_id[i]) + "',1,'" + str(name[i]) + "','" + str(description[i]) + "','" + str(meta_title[i]) + "','" + str(meta_description[i]) + "','" + str(meta_keyword[i]) + "'),\n"

    # oc_product_to_store
    oc_product_to_store += "('" + str(product_id[i]) + "',0),\n"

    # oc_product_attribute
    if (not isinstance(attribute_ids[i],float)):
        attributes = attribute_ids[i].split(",")
        text = texts[i].split(",")
        for j in range(len(attributes)):
            oc_product_attribute += "('" + str(product_id[i]) + "','" + str(attributes[j]) + "',1,'" + str(text[j]) + "'),\n"
    else:
        oc_product_attribute += "('" + str(product_id[i]) + "','" + str(attribute_ids[i]) + "',1,'" + str(texts[i]) + "'),\n"
    

    # oc_product_option_value
    if (not isinstance(quantities[i],float)):
        units2 = units[i].split(",")
        quantities2 = quantities[i].split(",")
        prices2 = prices[i].split(";")
        option_value_id2 = option_value_id[i].split(",")
        for j in range(len(prices2)):
            oc_product_option_value += "('" + str(product_option_value_id) + "','" + str(product_option_id[i]) + "',13,'" + str(product_id[i]) + "','" + str(option_value_id2[j]) + "','" + str(units2[j]) + "','" + str(quantities2[j]) + "','" + str(prices2[j]) + "'),\n"
            # increments product_option_value_id
            product_option_value_id += 1
    else:
        oc_product_option_value += "('" + str(product_option_value_id) + "','" + str(product_option_id[i]) + "',13,'" + str(product_id[i]) + "','" + str(option_value_id[i]) + "','" + str(units[i]) + "','" + str(quantities[i]) + "','" + str(prices[i]) + "'),\n"
        # increments product_option_value_id
        product_option_value_id += 1

    # oc_product_option
    oc_product_option += "('" + str(product_id[i]) + "',13),\n"
    

    # oc_product_discount

    # oc_product_special

    # oc_product_image
    oc_product_image += "('" + str(product_id[i]) + "','catalog/" + str(image[i]) + "'),\n"

    # oc_product_to_category
    # oc_product_to_category2
    if (not isinstance(categories[i],float)):
        category = categories[i].split(",")
        for j in range(len(category)):
            oc_product_to_category += "('" + str(product_id[i]) + "','" + str(category[j]) + "'),\n"
            oc_product_to_category2 += "('" + str(product_id[i]) + "','" + str(category[j]) + "'),\n"
    else:
        oc_product_to_category += "('" + str(product_id[i]) + "','" + str(categories[i]) + "'),\n"
        oc_product_to_category2 += "('" + str(product_id[i]) + "','" + str(categories[i]) + "'),\n"

    # oc_product_to_layout
    oc_product_to_layout += "('" + str(product_id[i]) + "'),\n"
    

    # oc_url_alias
    oc_url_alias += "('product_id=" + str(product_id[i]) + "','" + str(SEO[i]) + "'),\n"

    


# after for loop
# last iteration
i = i + 1

# delete
delete_oc_product += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_description += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_to_store += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_attribute += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_image += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_to_category += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_to_category2 += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_to_layout += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_url_alias += "query='product_id=" + str(product_id[i]) + "';\n"

delete_oc_product_option += "product_id='" + str(product_id[i]) + "';\n"
delete_oc_product_option_value += "product_id='" + str(product_id[i]) + "';\n"
                                        
# oc_product
oc_product += "('" + str(product_id[i]) + "','Molecules',999,6,'catalog/" + str(image[i]) + "',1,'" + datetime.datetime.today().strftime('%Y-%m-%d') + "',1,1,1,1,1,NOW(),NOW(),'" + str(library[i]) + "','" + str(library_base_price[i]) + "');\n\n"

# oc_product_description
oc_product_description += "('" + str(product_id[i]) + "',1,'" + str(name[i]) + "','" + str(description[i]) + "','" + str(meta_title[i]) + "','" + str(meta_description[i]) + "','" + str(meta_keyword[i]) + "');\n\n"

# oc_product_to_store
oc_product_to_store += "('" + str(product_id[i]) + "',0);\n\n"

# oc_product_attribute
if (not isinstance(attribute_ids[i],float)):
    attributes = attribute_ids[i].split(",")
    text = texts[i].split(",")
    for j in range(len(text)-1):
        oc_product_attribute += "('" + str(product_id[i]) + "','" + str(attributes[j]) + "',1,'" + str(text[j]) + "'),\n"
        
    j = j + 1
    oc_product_attribute += "('" + str(product_id[i]) + "','" + str(attributes[j]) + "',1,'" + str(text[j]) + "');\n\n"
else:
    oc_product_attribute += "('" + str(product_id[i]) + "','" + str(attribute_ids[i]) + "',1,'" + str(texts[i]) + "');\n\n"

# oc_product_option_value
if (not isinstance(quantities[i],float)):
    units2 = units[i].split(",")
    quantities2 = quantities[i].split(",")
    prices2 = prices[i].split(";")
    option_value_id2 = option_value_id[i].split(",")
    for j in range(len(prices2)-1):
        oc_product_option_value += "('" + str(product_option_value_id) + "','" + str(product_option_id[i]) + "',13,'" + str(product_id[i]) + "','" + str(option_value_id2[j]) + "','" + str(units2[j]) + "','" + str(quantities2[j]) + "','" + str(prices2[j]) + "'),\n"
        # increments product_option_value_id
        product_option_value_id += 1
    j += 1
    oc_product_option_value += "('" + str(product_option_value_id) + "','" + str(product_option_id[i]) + "',13,'" + str(product_id[i]) + "','" + str(option_value_id2[j]) + "','" + str(units2[j]) + "','" + str(quantities2[j]) + "','" + str(prices2[j]) + "');\n\n"
else:
    oc_product_option_value += "('" + str(product_option_value_id) + "','" + str(product_option_id[i]) + "',13,'" + str(product_id[i]) + "','" +str(option_value_id[i]) + "','" + str(units[i]) + "','" + str(quantities[i]) + "','" + str(prices[i]) + "');\n\n"

# oc_product_option
oc_product_option += "('" + str(product_id[i]) + "',13);\n\n"


# oc_product_image
oc_product_image += "('" + str(product_id[i]) + "','catalog/" + str(image[i]) + "');\n\n"

# oc_product_to_category
# oc_product_to_category2
if (not isinstance(categories[i],float)):
    category = categories[i].split(",")
    for j in range(len(category)-1):
        oc_product_to_category += "('" + str(product_id[i]) + "','" + str(category[j]) + "'),\n"
        oc_product_to_category2 += "('" + str(product_id[i]) + "','" + str(category[j]) + "'),\n"

    j = j + 1
    oc_product_to_category += "('" + str(product_id[i]) + "','" + str(category[j]) + "');\n\n"
    oc_product_to_category2 += "('" + str(product_id[i]) + "','" + str(category[j]) + "');\n\n"
else:
    oc_product_to_category += "('" + str(product_id[i]) + "','" + str(categories[i]) + "');\n\n"
    oc_product_to_category2 += "('" + str(product_id[i]) + "','" + str(categories[i]) + "');\n\n"
# oc_product_to_layout
oc_product_to_layout += "('" + str(product_id[i]) + "');\n\n"

# oc_url_alias
oc_url_alias += "('product_id=" + str(product_id[i]) + "','" + str(SEO[i]) + "');\n\n"

with codecs.open("../productImportSQL/" + file_name,"w", "utf-8-sig") as file:
    # writing
    # delete
    file.write(delete_oc_product)
    file.write(delete_oc_product_description)
    file.write(delete_oc_product_to_store)
    file.write(delete_oc_product_attribute)
    file.write(delete_oc_product_option)
    file.write(delete_oc_product_option_value)
    #file.write(delete_oc_product_discount)
    #file.write(delete_oc_product_special)
    file.write(delete_oc_product_image)
    file.write(delete_oc_product_to_category)
    file.write(delete_oc_product_to_category2)
    #file.write(delete_oc_product_filter)
    #file.write(delete_oc_product_related)
    #file.write(delete_oc_product_reward)
    file.write(delete_oc_product_to_layout)
    #file.write(delete_oc_product_recurring)
    file.write(delete_oc_url_alias)

    file.write("\n")
    # insertions
    file.write(oc_product)
    file.write(oc_product_description)
    file.write(oc_product_to_store)
    file.write(oc_product_attribute)
    file.write(oc_product_option)
    file.write(oc_product_option_value)
    #file.write(oc_product_option)
    #file.write(oc_product_discount)
    #file.write(oc_product_special)
    file.write(oc_product_image)
    file.write(oc_product_to_category)
    file.write(oc_product_to_category2)
    #file.write(oc_product_filter)
    #file.write(oc_product_related)
    #file.write(oc_product_reward)
    file.write(oc_product_to_layout)
    #file.write(oc_product_recurring)
    file.write(oc_url_alias)


file.close()

print(file_name + " saved.")









