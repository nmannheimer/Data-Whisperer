import xml.etree.ElementTree as ET
import os
import time
from tkFileDialog import askopenfilename
# from difflib import SequenceMatcher as SM
# Not currently implementing fuzzy matching with difflib
# This would allow the watching of close text matches so 'Sals' == 'Sales'


###
# TO DO - List of top features to implement
###

# In general - clean and simplify code

# 1) Support for other chart types like maps and scatter plots
# 2) Upload to Server using Extract API
# 3) This is the big one, refactor the dialogue tree to be sheet focused and let you select a sheet to work on
#        + Modify existing sheets (display a list of existing sheets and choose one)
#        + Add ability to use swap axis
#        + Ability to name sheets
# 4) Multiple dimensions and measures on rows and columns
# 5) Handle date aggregations on color
# 6) Explore adding more data sources beyond SQL and Excel
# 7) Add more aggregation types and ability to sort results
# 8) Support for multiple data sources

###
# THIS IS WHERE THE FUNCTIONS THAT MODIFY XML ARE
###

# Initial XML Schema Read
def load_data(source):
    tree = ET.parse(source)
    root = tree.getroot()
    dimensions = []
    measures = []
    metadata = {}

    # Decide whether source is Excel or SQL
    # Then Parse XML to find Dimensions and Measure in the Data
    if root.findall('.//named-connection/connection')[0].attrib['class'] == 'excel-direct':
        for column in root.findall('.//columns/column'):
            metadata[column.attrib['name']] = column.attrib['datatype']
        for key, value in metadata.iteritems():
            if value in ['integer', 'real']:
                measures.append(key)
            elif value in ['date', 'string']:
                dimensions.append(key)
    elif root.findall('.//named-connection/connection')[0].attrib['class'] == 'sqlserver':
        for field in root.iter('metadata-record'):
            metadata[field.find('remote-name').text] = field.find('local-type').text
        for key, value in metadata.iteritems():
            if value in ['integer', 'real']:
                measures.append(key)
            elif value in ['date', 'string']:
                dimensions.append(key)

    # Find data source name
    dsourceloc = root.findall('.//datasource')
    dsource = []
    for source in dsourceloc:
        if source.attrib['name'] != 'Parameters':
            dsource = source.attrib['name']

    return root, tree, dimensions, measures, metadata, dsource


# Create a new Worksheet
def create_worksheet(count, root, workbook, output, dsource, worksheet_name):

    # We need to handle the fact that if two sheets have the same name, the workbook will crash
    # We do this by inserting the loop counter number into the sheet name so no sheets will be identical
    if count == 0:
        worksheet_name = '1) ' + worksheet_name
    else:
        worksheet_name = str(count + 1) + ') ' + worksheet_name

    # build the various components of a standard worksheet
    if count == 0:
        for worksheet in root.iter('worksheets'):
            firstsheet = worksheet.find('worksheet[1]')
            worksheet.remove(firstsheet)
    insertpoint = root.find('./worksheets')
    newsheet = ET.SubElement(insertpoint,'worksheet')
    newsheet.set('name', worksheet_name)
    table = ET.SubElement(newsheet, 'table')
    view = ET.SubElement(table, 'view')
    style = ET.SubElement(table, 'style')
    panes = ET.SubElement(table, 'panes')
    datasources = ET.SubElement(view,'datasources')
    datasource = ET.SubElement(datasources,'datasource').set('name', dsource)
    datasource_dependencies = ET.SubElement(view, 'datasource-dependencies').set('datasource', dsource)
    aggregation = ET.SubElement(view, 'aggregation').set('value', 'true')
    pane = ET.SubElement(panes, 'pane')
    pane_view = ET.SubElement(pane, 'view')
    breakdown = ET.SubElement(pane_view, 'breakdown').set('value', 'auto')
    encodings = ET.SubElement(pane, 'encodings')

    # encodings needs to have <></> in order to function or it will throw an error
    # this function creates a sub-element and then removes it to cause this
    dummy_encoding = ET.SubElement(encodings, 'fake')
    encodings.remove(dummy_encoding)

    # Let's add those great Worksheet cards
    # For all sheets except for the first one (which will already have them)...
    # Then workbook will crash if there are two identical sheet names, so handle this by including an increasing number
    # This prevents the error where running the same query will brick the workbook
    if count == 0:
        window = root.find('.//windows/window')
        window.set('name', worksheet_name)

    if count != 0:
        # Insert the highest level cards that go with a sheet
        cardinsert = root.find('.//windows')
        window = ET.SubElement(cardinsert, 'window')
        cards = ET.SubElement(window, 'cards')
        window.set('class', 'worksheet')
        window.set('maximized', 'true')
        window.set('name', worksheet_name)
        cards.set('name', 'right')

        # Insert the default left side cards that should go with every sheet
        cardinsert = root.findall('.//window[{0}]/cards'.format(count+1))
        edge_left = ET.SubElement(cardinsert[0], 'edge')
        strip = ET.SubElement(edge_left, 'strip')
        card1 = ET.SubElement(strip, 'card')
        card2 = ET.SubElement(strip, 'card')
        card3 = ET.SubElement(strip, 'card')
        edge_left.set('name', 'left')
        strip.set('size', '160')
        card1.set('type', 'pages')
        card2.set('type', 'filters')
        card3.set('type', 'marks')

        # Insert the default top cards that should go with every sheet
        edge_top = ET.SubElement(cardinsert[0], 'edge')
        strip1 = ET.SubElement(edge_top, 'strip')
        strip2 = ET.SubElement(edge_top, 'strip')
        strip3 = ET.SubElement(edge_top, 'strip')
        card1 = ET.SubElement(strip1, 'card')
        card2 = ET.SubElement(strip2, 'card')
        card3 = ET.SubElement(strip3, 'card')
        edge_top.set('name', 'top')
        strip1.set('size', '2147483647')
        strip2.set('size', '2147483647')
        strip3.set('size', '2147483647')
        card1.set('type', 'columns')
        card2.set('type', 'rows')
        card3.set('type', 'title')

# Write the XML to the workbook
    tree.write('{0}{1}.twb'.format(output, workbook))


# Function to insert the row and column values for a viz
# Currently only designed to take one dimension on columns and one measure on rows
def row_column(aggregation, row, column, root, tree, count, dsource, workbook, date, output, cont_disc, date_agg, metadata):

    # If a row is present, insert it
    if row:
        # Insert the row/measure in the datasource-dependencies "column-instance"
        insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
        newdata = ET.SubElement(insertpoint[0], 'column-instance')
        newdata.set('column', '['+row+']')
        newdata.set('derivation', aggregation)
        newdata.set('type', "quantitative")
        newdata.set('pivot', "key")
        newdata.set('pivot', '['+aggregation.lower()+':'+row+':qk]')

        # Insert the row/measure in the datasource-dependencies "column"
        insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
        newdata = ET.SubElement(insertpoint[0], 'column')
        newdata.set('datatype', metadata[row])
        newdata.set('name', '['+row+']')
        newdata.set('role', 'measure')
        newdata.set('type', 'quantitative')

        # Insert the row/measure into the work sheet rows section
        insertpoint = root.findall('.//worksheet[{0}]/table'.format(count+1))
        newrow = ET.SubElement(insertpoint[0], 'rows')
        newrow.text = '['+dsource+'].['+aggregation.lower()+':'+row+':qk]'

    # Dates are handled differently since they must be aggregated (normal dimensions do not)
    if column and date == 1:
        state = ''
        agg = ''
        type = ''
        # Differentiate needed fields if the date is continuous or discrete
        if cont_disc == 'Continuous':
            state = 'qk'
            aggs = {'Year': 'tyr', 'Quarter': 'tqr', 'Month':'tmn', 'Week':'twk', 'Day':'tdy'}
            agg = aggs[date_agg]
            type = 'quantitative'
            date_agg += '-Trunc'
        elif cont_disc == 'Discrete':
            state = 'ok'
            aggs = {'Year': 'yr', 'Quarter': 'qr', 'Month':'mn', 'Week':'wk', 'Day':'dy'}
            agg = aggs[date_agg]
            type = 'ordinal'

        # Insert the column-instance (needed if we aggregate)
        insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
        newdata = ET.SubElement(insertpoint[0], 'column-instance')
        newdata.set('column', '['+column+']')
        newdata.set('derivation', date_agg)
        newdata.set('type', type)
        newdata.set('pivot', "key")
        newdata.set('name', '['+agg+':'+column+':'+state+']')

        # Insert the column
        insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
        newdata = ET.SubElement(insertpoint[0], 'column')
        newdata.set('datatype', metadata[column])
        newdata.set('name', '['+column+']')
        newdata.set('role', 'dimension')
        newdata.set('type', type)

        # Insert the dimension onto columns
        insertpoint = root.findall('.//worksheet[{0}]/table'.format(count+1))
        newcol = ET.SubElement(insertpoint[0], 'cols')
        newcol.text = '['+dsource+'].['+agg+':'+column+':'+state+']'

    # Standard non-aggregate dimension
    elif column:
        # Insert the dimension onto columns
        insertpoint = root.findall('.//worksheet[{0}]/table'.format(count+1))
        newcol = ET.SubElement(insertpoint[0], 'cols')
        newcol.text = '['+dsource+'].[none:'+column+':nk]'

        # Insert the non-aggregated dimension onto columns in datasource dependencies
        insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
        newdata = ET.SubElement(insertpoint[0], 'column')
        newdata.set('datatype', metadata[column])
        newdata.set('name', '['+column+']')
        newdata.set('role', 'dimension')
        newdata.set('type', 'ordinal')

    tree.write('{0}{1}.twb'.format(output, workbook))


# Change the Mark type
def change_mark(newmark, root, tree, count, workbook, output):

    # Insert the selected chart type
    insertpoint = root.findall('.//worksheet[{0}]/table/panes/pane'.format(count+1))
    insertmark = ET.SubElement(insertpoint[0], 'mark')
    insertmark.set('class', newmark.replace(" ", ""))
    tree.write('{0}{1}.twb'.format(output, workbook))


# Place a new field on Color (takes a dimension or aggregated measure)
def change_color(color, count, dsource, workbook, col_date, output, field_type, col_agg):

    # Insert Color
    insertpoint = root.findall('.//worksheet[{0}]/table/panes/pane/encodings'.format(count+1))
    insertcolor = ET.SubElement(insertpoint[0], 'color')

    # Develop color legend card
    cardinsert = root.findall('.//window[{0}]/cards'.format(count+1))
    edge = ET.SubElement(cardinsert[0], 'edge')
    strip = ET.SubElement(edge, 'strip')
    card = ET.SubElement(strip, 'card')
    edge.set('name', 'right')
    strip.set('size', '160')
    card.set('pane-specification-id', '0')
    card.set('type', 'color')

    # Choose different XML insert depending on whether field_type=1 (dimension) or 0 (measure)
    if field_type == 1:
        # Handle aggregation if color field is a date
        # We need to add a column-instance field if there is a date on color and column (otherwise workbook crashes)
        if col_date == 1:
            # Set the color
            insertcolor.set('column', '['+dsource+'].[yr:'+color+':ok]')

            # Add the column instance
            insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
            newdata = ET.SubElement(insertpoint[0], 'column-instance')
            newdata.set('column', '['+color+']')
            newdata.set('derivation', 'Year')
            newdata.set('name', '[yr:'+color+':ok]')
            newdata.set('pivot', "key")
            newdata.set('type', 'ordinal')

        else:
            # Set the color for a regular dimension
            insertcolor.set('column', '['+dsource+'].[none:'+color+':nk]')

    elif field_type == 0:
            # Set the color
            insertcolor.set('column', '['+dsource+'].['+col_agg.lower()+':'+color+':qk]')

            # Set the column-instance for an aggregate measure
            insertpoint = root.findall('.//worksheet[{0}]/table/view/datasource-dependencies'.format(count+1))
            newdata = ET.SubElement(insertpoint[0], 'column-instance')
            newdata.set('column', '['+color+']')
            newdata.set('derivation', col_agg)
            newdata.set('name', '['+col_agg.lower()+':'+color+':qk]')
            newdata.set('pivot', "key")
            newdata.set('type', 'quantitative')

    tree.write('{0}{1}.twb'.format(output, workbook))


# Place a dimension on Detail
def change_detail(detail, count, dsource, workbook, det_date, output):

    # Write the XML for a detail
    insertpoint = root.findall('.//worksheet[{0}]/table/panes/pane/encodings'.format(count+1))
    insertlod = ET.SubElement(insertpoint[0], 'lod')

    # Select different aggregation for a date dimension vs a standard dimension
    if det_date == 1:
        insertlod.set('column', '['+dsource+'].[yr:'+detail+':ok]')
    else:
        insertlod.set('column', '['+dsource+'].[none:'+detail+':nk]')

    tree.write('{0}{1}.twb'.format(output, workbook))


# Swap row and column values
# The default will place the dimension on columns, this allows the dimension to be placed on rows
def swap_axis(workbook, output, count):
    oldcol = []
    oldrow = []
    for rows in root.findall('.//worksheet[{0}]/table/rows'.format(count+1)):
        oldrow = rows.text
    for columns in root.findall('.//worksheet[{0}]/table/cols'.format(count+1)):
        oldcol = columns.text
    for rows in root.findall('.//worksheet[{0}]/table/rows'.format(count+1)):
        rows.text = oldcol
    for columns in root.findall('.//worksheet[{0}]/table/cols'.format(count+1)):
        columns.text = oldrow
    tree.write('{0}{1}.twb'.format(output, workbook))


# Print the measures and dimensions available in the workbook/dataset
def print_dims_meas(dimensions, measures):
    print "The data set contains the following measures:\n"
    for meas in measures:
        print meas
    print ""
    print "The data set contains the following dimensions:\n"
    for dim in dimensions:
        print dim
    print ""

# Print the dimensions available in the workbook/dataset
def print_dims(dimensions):
    print ""
    print "The data set contains the following dimensions:\n"
    for dim in dimensions:
        print dim
    print ""

# Print the measures available in the workbook/dataset
def print_meas(measures):
    print ""
    print "The data set contains the following measures:\n"
    for meas in measures:
        print meas
    print ""

# Open the .twb file and render the view
def render(workbook, output):
    os.system('start ' + '{0}{1}.twb'.format(output, workbook))


# Print the list of available inputs into the command line
def commands():
    print ""
    print "Commands:\n" \
        "    'Show me profit for each state'/'Let me see average sales by region'/'Profit by segment as an area chart' - Enter a query to render\n" \
        "    Data - Show available Dimensions and Measures\n" \
        "    Commands or Help - Show available Commands\n" \
        "    Aggregations - Show available measure aggregations\n" \
        "    Charts - Show available chart types\n" \
        "    Exit - Leave Data Whisperer\n"


# Print the list of available aggregations
def aggregations():
    print ""
    print "Aggregations:\n" \
        "    'Sum', 'Average', 'Min', 'Max', 'Count'"


# Print the list of available aggregations
def chart_types():
    print ""
    print "Chart Types:\n" \
        "    'Line', 'Square', 'Gantt Bar', 'Bar', 'Circle', 'Area'"

###
# THIS IS WHERE THE USER INTERACTION AND QUERY PARSING PART OF THE PROGRAM IS
###

print"""
-----------------------------------------------------------------------------------
______  ___ _____ ___    _    _ _   _ _____ ___________ ___________ ___________
|  _  \/ _ |_   _/ _ \  | |  | | | | |_   _/  ___| ___ |  ___| ___ |  ___| ___ |
| | | / /_\ \| |/ /_\ \ | |  | | |_| | | | \ `--.| |_/ | |__ | |_/ | |__ | |_/ /
| | | |  _  || ||  _  | | |/\| |  _  | | |  `--. |  __/|  __||    /|  __||    /
| |/ /| | | || || | | | \  /\  | | | |_| |_/\__/ | |   | |___| |\ \| |___| |\  |
|___/ \_| |_/\_/\_| |_/  \/  \/\_| |_/\___/\____/\_|   \____/\_| \_\____/\_| \_|

-----------------------------------------------------------------------------------
"""

# load the XML from the workbook and return the key features of the XML
print "Welcome to Data Whisperer, the Text to Viz CMD Line interface for Tableau!"
time.sleep(2)
print "To get started, we'll need to locate a .twb workbook that is connected to a SINGLE SQL or Excel data source.\n"
time.sleep(2)
print "Let's find a source .twb file..."
time.sleep(2)
print "Navigate to a .twb file in the pop-up window."
time.sleep(3)
source = askopenfilename()
root, tree, dimensions, measures, metadata, dsource = load_data(source)
print "Great, we've loaded {0} into memory!".format(source)
print ""

# Choose a location to output the new XML/TWB file to
output = 'C:/Users/Public/Desktop/'
new_output = raw_input("Enter file path for desired save location (press Enter to default to Desktop):\n")
if new_output != "":
    if not new_output.endswith("/"):
        new_output += "/"
    output = new_output.replace("\\", "/")
print "Output location = {0}\n".format(output)

# Choose a name for the new workbook and remove any forbidden characters for windows file names
workbook = raw_input("What would you like to call your new workbook?\n").replace(" ", "")
while True:
    if len(workbook) == 0:
        workbook = raw_input("Please enter a valid workbook name longer than 0 characters.\n").replace(" ", "")
        continue
    elif any(char in workbook for char in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']):
        workbook = raw_input("Please enter a valid workbook file name not including forbidden characters.\n").replace(" ", "")
        continue
    else:
        break

print ""

# Display available dimensions and measures in the workbook to the user
print_dims_meas(dimensions, measures)

query = ""
# Count will count the number of times we go through the loop
# This is important for locating and naming sheets
count = 0

# This is the main chat loop that allows the repeated generation of sheets
while True:

    # Take the user input for what they would like to see
    if count < 1:                                                   # After the first loop, display the new prompt
        commands()
        query = raw_input("How would you like to see your data? (For example: 'Show me average quantity by region')\n")
        if query == "":
            query = raw_input("Please enter a valid query (For example: 'Show me average quantity by region')\n")
    else:
        query = raw_input("What would you like to see next?\n")
        if query == "":
            query = raw_input("Please enter a valid query (For example: 'Show me average quantity by region')\n")

    # Exit the program if the user enters exit
    if query.title() == 'Exit':
        print "Thank you for using Data Whisperer!"
        break

    # Print the available data in the workbook
    if query.title() == 'Data':
        print_dims_meas(dimensions, measures)
        continue

    # Re-Print the available commands for Data Whisperer
    if count == 0 and (query.title() == 'Commands' or query.title() == 'Help'):
        continue
    elif query.title() == 'Commands' or query.title() == 'Help':
        commands()
        continue

    # Re-Print the available aggregations for Data Whisperer
    if query.title() == 'Aggregations':
        aggregations()
        continue

    # Re-Print the available chart types for Data Whisperer
    if query.title() == 'Charts':
        chart_types()
        continue

    # Clean up the user input by removing punctuation and capitalizing all words
    query = query.strip(".").strip("?").strip("!").strip(",").strip('\\').title()
    measure = ""
    dimension = ""
    chart = ""
    aggregation = ""

    # Parse the user input by checking whether each measure, dimension, and chart type is found in the query
    # Return a Measure, Dimension, Aggregation, and Chart type from the query
    for value in measures:
        if value.title() in query:
            measure = value
    for value in dimensions:
        if value.title() in query:
            dimension = value
    for agg in ['Sum', 'Min', 'Max', 'Average', 'Count']:
        if agg in query:
            aggregation = agg
            break
    if measure == "" and dimension == "":
        print ""
        print "Please Enter a valid dimension or measure."
        continue

    # If no aggregation is specified, default to 'Sum' and handle Tableau needing 'avg'
    if aggregation == "":
        aggregation = 'Sum'
    elif aggregation == "Average":
        aggregation = 'Avg'
    elif aggregation == 'Count':
        aggregation = 'Cnt'

    # Parse the query to see if the user asked for a specific chart type
    for chart_type in ['Line', 'Square', 'Gantt Bar', 'Bar', 'Circle', 'Area']:
        if chart_type in query:
            chart = chart_type
            break
    # If no chart type is specified, default to 'Automatic'
    if chart == "":
        chart = 'Automatic'

    # Handle a date needing special aggregation for a dimension
    cont_disc = ""
    date_agg = ""
    date = 0
    # Give additional aggregation options if a date dimension is included in the query
    if dimension and metadata[dimension] == 'date':
        date = 1
        cont_disc = raw_input("It looks like " + dimension + " is a date. Would you like to see it as a continuous or discrete date?\n").title()
        # Choose continuous or discrete date
        while cont_disc not in ['Continuous', 'Discrete']:
            cont_disc = raw_input("Please choose Continuous or Discrete:\n").title()
        # Choose a date aggregation
        date_agg = raw_input("How would you like to aggregate the date? ie 'Year' or 'Month'?\n").title()
        while date_agg not in ['Year', 'Quarter', 'Month', 'Week', 'Day']:
            date_agg = raw_input("Please choose Year, Quarter, Month, Week or Day:\n").title()

    # Name the new sheet and handle possibly having no dimension or measure
    if measure != "" and dimension != "":
        worksheet_name = aggregation + " " + measure + " " + "by" + " " + dimension
    elif measure != "" and dimension == "":
        worksheet_name = aggregation + " " + measure
    elif measure == "" and dimension != "":
        worksheet_name = dimension
    # Shouldn't be possible to hit this but, eh
    else:
        worksheet_name = "Sheet" + " " + str(count+1)

    # Build a new worksheet
    create_worksheet(count, root, workbook, output, dsource, worksheet_name)

    # Load parsed variables into the functions to modify the XML for a dimension and measure and mark type
    row_column(aggregation, measure, dimension, root, tree, count, dsource, workbook, date, output, cont_disc, date_agg, metadata)
    change_mark(chart, root, tree, count, workbook, output)

    # If there is a dimension field in the query, the user chooses how they want to orient the view
    if dimension != "":
        swap = ''
        while any(angle not in swap for angle in ['Columns', 'Column', 'Rows', 'Row']):
            swap = raw_input("Would you like the dimensional field, " + dimension + ", to be on 'Rows' or 'Columns'?\n").title()
            if 'Columns' in swap or 'Column' in swap:
                # Do Nothing
                break
            elif 'Rows' in swap or 'Row' in swap:
                # Swap the axes to put the dimension on Rows
                swap_axis(workbook, output, count)
                break
            elif 'No' in swap or 'Exit' in swap:
                break
            else:
                continue

    # Offer the user the option to add a dimension or aggregate measure to color
    # We also need to again test for dates, which require special aggregation
    col_date = 0
    det_date = 0
    color_query = ""
    color = ""
    col_agg = 'Sum'
    # This section will handle a dimension or measure being placed on color, cleaning and parsing the input
    if measure != "" or dimension != "":
        # Take a new input for visualization color and clean it up
        color_query = raw_input("Would you like to place a dimension or measure on color? (Press Enter to skip)\n")
        color_query = color_query.strip(".").strip("?").strip("!").strip(",").title()
        # Move on if there is no query or if the user enters 'No' or 'Exit'
        if color_query == "" or color_query == "No" or color_query == 'Exit':
            pass
        # Using a method similar to the rows columns function, clean and identify a dimension or measure and aggregation
        else:
            while color == "":
                # Identify the dimension or measure
                for value in dimensions + measures:
                    if value.title() in color_query:
                        color = value
                        break
                # Identify the aggregation
                for agg in ['Sum', 'Min', 'Max', 'Average', 'Count']:
                    if agg in color_query:
                        col_agg = agg
                        if col_agg == 'Average':
                            col_agg = 'Avg'
                        elif col_agg == 'Count':
                            col_agg = 'Cnt'
                        break
                    else:
                        col_agg = 'Sum'
                # If we didn't detect a color ask again or give the option to skip
                if color == "":
                    color_query = raw_input("Please select a valid dimension or measure or enter 'Exit' to move on:\n")
                    color_query = color_query.strip(".").strip("?").strip("!").strip(",").title()
                    if color_query == 'Exit':
                        break
                    else:
                        continue
                # If we do have a color, detect if it's a dimension or measure and if it's a date, then write to XML
                if color != "":
                    if metadata[color] == 'date':
                        col_date = 1
                    if color in dimensions:
                        field_type = 1
                    else:
                        field_type = 0
                    change_color(color, count, dsource, workbook, col_date, output, field_type, col_agg)

        # Offer the user the option to add a field to LOD
        # Currently only dimensions are supported for LOD fields
        detail = ''
        detail_query = raw_input("Would you like to place a dimension on detail? (Press Enter to skip)\n").title()
        if not (detail_query == "" or detail_query == "No" or detail_query == "Exit"):
            while detail == '':
                if detail_query == 'Exit':
                    break
                for value in dimensions:
                    if value.title() in detail_query:
                        detail = value
                        break
                if detail == '':
                    detail_query = raw_input("Please select a valid dimension or enter 'Exit' to move on:\n").title()
            # Dates are supported for LODs
            if detail != '':
                if metadata[detail] == 'date':
                    det_date = 1
                change_detail(detail, count, dsource, workbook, det_date, output)

    # Open the modified workbook in Tableau
    render(workbook, output)
    count += 1

    # Show this message so it looks like something is happening while Tableau is loading...
    print "Rendering."
    time.sleep(1)
    print "Rendering.."
    time.sleep(1)
    print "Rendering..."

# End of functional code
