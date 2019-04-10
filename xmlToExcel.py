from xml.dom import minidom
import xlsxwriter
import os


def translate(project_path, excel_path):
    print(project_path)
    print(excel_path)
    list_dictionary = {}

    translation_list = []
    string_path_list = []

    # Get list of all directories and files - strings.xml and values from the project path
    for subdir, dirs, files in os.walk(project_path):
        for file in files:
            if file.startswith('strings') and file.endswith('.xml'):
                string_path_list.append(subdir + '/' + file)
                trans_end = subdir[subdir.find('values') + 7:]
                if trans_end == '':
                    trans_end = 'en'
                translation_list.append(trans_end)
                continue
    print(translation_list)

    # add all strings to dictionary
    for i in range(len(string_path_list)):
        xml_doc = minidom.parse(string_path_list[i])
        item_list = xml_doc.getElementsByTagName('string')
        for item in item_list:
            if len(item.childNodes) > 0:
                key = item.attributes['name'].value
                value = item.childNodes[0].nodeValue

                if key not in list_dictionary:
                    list_dictionary[key] = {}
                list_dictionary[key][translation_list[i]] = value

    # create workbook
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet()

    # Change Cell color when empty
    red_format = workbook.add_format({'bg_color': '#d61111'})
    row = 0
    col = 0
    worksheet.write(row, col, 'Key')

    for t in translation_list:
        col += 1
        worksheet.write(row, col, t)

    row += 1

    # Insert the values with respect to the key of each translation
    for key, value in list_dictionary.items():
        col = 0
        worksheet.write(row, col, key)

        for t in translation_list:
            col += 1

            if t in value:
                worksheet.write(row, col, value[t])

            else:
                worksheet.write(row, col, " ", red_format)

        row += 1

    workbook.close()


def main():
    project_path = input("Enter Project Path up un-till src/main/res : ")
    excel_path = input("Enter Excel sheet Path (Create the sheet first using Excel) : ")
    translate(project_path, excel_path)


if __name__ == "__main__":
    main()
