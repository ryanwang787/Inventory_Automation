import excelrd
import xlsxwriter


def write_sheet(loc, name, title_list, title_width):
    """
    Function that writes all the rows that found a keyword into a new spreadsheet
    :param loc: Name of spreadsheet you want to search through
    :param name: What you want to name your new sorted spreadsheet
    :param title_list: The column labels for the original spreadsheet
    :param title_width: The column widths for the original spreadsheet
    :return: None. Creates a new spreadsheet of the results of the keyword search
    """
    wb = excelrd.open_workbook(loc)  # reading workbook
    sheet = wb.sheet_by_index(1)  # sheet index
    row_list = []  # init list of rows

    # adding each row of the spreadsheet minus the headers to our list
    for i in range(1, sheet.nrows):
        row_list.append(sheet.row_values(i))

    # loop that removes all extra whitespaces
    for i in range(len(row_list)):
        for j in range(len(row_list[i])):
            if type(row_list[i][j]) is str:  # check if the element is a string
                row_list[i][j] = ' '.join(row_list[i][j].split())  # join the string by single spaces

    workbook = xlsxwriter.Workbook(name)  # writing new workbook
    worksheet = workbook.add_worksheet()  # adding spreadsheet
    bold = workbook.add_format({'bold': True})  # bold formatting

    formatting(worksheet, title_list, title_width, bold)
    search(row_list, sheet, worksheet)

    workbook.close()


def formatting(worksheet, title_list, title_width, bold):
    """
    Function that formats the new spreadsheet the same as the original one
    :param worksheet: New sheet to be created
    :param title_list: List of the column titles for the original spreadsheet
    :param title_width: List of the column widths for the original spreadsheet
    :param bold: Bold formatting
    :return: None
    """

    # adding the keyword location columns
    title_list.append('Keyword location')
    title_width.append(18)
    # setting headers and cell widths
    for i in range(len(title_list)):
        worksheet.set_column(i, i, title_width[i])
        worksheet.write(0, i, title_list[i], bold)


def search(list, sheet, worksheet):
    """
    Function that searches through the cells of the original spreadsheet for the keywords
    :param list: Rows of the original spreadsheet
    :param sheet: Original sheet
    :param worksheet: New spreadsheet
    :return: None. Writes the rows that found a keyword to the new sheet, including which keyword was found
    """

    row = 1  # init row
    # run through the list of keywords and check if the keywords appear in our rows
    for i in list:
        print(row)
        print(i[-2])
        SDG_location = []
        keyword = i[-2].split(', ')
        print(keyword)
        for j in range(len(keyword)):
            # checking if keyword is in title, or description
            if keyword[j] in str(i[1]).lower() and keyword[j] in str(i[2]).lower():
                SDG_location.append('Title and description')
            elif keyword[j] in str(i[1]).lower() and keyword[j] not in str(i[2]).lower():
                SDG_location.append('Title')
            elif keyword[j] in str(i[2]).lower() and keyword[j] not in str(i[1]).lower():
                SDG_location.append('Description')

        print(SDG_location)
        if 'Title and description' in SDG_location:
            SDG_loc_str = 'Title and description'
        if 'Title' in SDG_location and 'Description' in SDG_location:
            SDG_loc_str = 'Title and description'
        if 'Title' in SDG_location and 'Description' not in SDG_location and 'Title and description' not in SDG_location:
            SDG_loc_str = 'Title'
        if 'Description' in SDG_location and 'Title' not in SDG_location and 'Title and description' not in SDG_location:
            SDG_loc_str = 'Description'

        for m in range(sheet.ncols):  # for every column in that row
            worksheet.write(row, m, i[m])  # write in the copied info for each cell

        worksheet.write(row, sheet.ncols, SDG_loc_str)  # add the keywords that were used
        row += 1



sheet_path = 'Course Inventory 2022-23.xlsx'  # name of full spreadsheet
wb_name = 'Test - SDG locations.xlsx'  # new workbook
# name

# headings; adjust as needed
titles = ['Course Code', 'Course Title', 'Course Description', 'Division', 'Unit', 'Keyword(s)', 'SDG(s)']
# width of headings; adjust as needed
col_width = [10, 50, 100, 30, 30, 18, 18]

# list of SDG keywords
keyword_list = ['poverty', 'income distribution', 'wealth distribution',
                'socio economic', 'socio-economic', 'socioeconomic', 'homeless',
                'low-income', 'low income', 'affordab', 'disparity', 'welfare',
                'social safety', 'developing country', 'vulnerability',
                'precarity', 'precarious', 'pro-poor', 'agricultur',
                'nutrition', 'food security', 'food insecurity', 'food-secure',
                'food system', 'child hunger', 'food justice', 'food scarcity',
                'food sovereignty', 'food culture', 'culinary', 'malnutrition',
                'agro', 'permaculture', 'indigenous crops',
                'regenerative agriculture', 'urban agriculture', 'organic food',
                'biodynamic', 'food literacy', 'food education',
                'benefit sharing', 'access and benefit sharing', 'ABS',
                'end hunger', 'food price', 'zero hunger', 'well being',
                'wellbeing', 'well-being', 'mental health', 'public health',
                'global health', 'health care', 'healthcare', 'health issues',
                'mental wellness', 'disabilit', 'sexual education',
                'mindfulness', 'holism', 'illness', 'health education',
                'communicable disease', 'health determinants', 'vaccine',
                'substance abuse', 'maternal mortality', 'family planning',
                'hazardous chemicals', 'pollution', 'health equity',
                'neonatal mortality', 'infant mortality', 'child health',
                'road traffic accidents', 'reproductive health', 'epidemics',
                'universal health coverage', 'equitable', 'pedagogy',
                'knowledge', 'worldview', 'learning', 'knowledges',
                'traditional knowledge', 'land-based knowledge',
                'place-based knowledge', 'decolonial', 'anticolonial',
                'settler', 'equitable', 'equity', 'anti-racism', 'racism',
                'anti-oppression', 'oppression,anti-discriminatory',
                'early childhood development', 'peace', 'citizen',
                'sustainability teaching', 'sustainability education',
                'universal literacy', 'basic literacy', 'universal numeracy',
                'environmental education',
                'education for sustainable development', 'ecojustice education',
                'eco justice education', 'place-based education',
                'humane education', 'land-based learning',
                'nature-based education', 'climate change education',
                'vocational', 'technical learning', 'free education',
                'accessible education', 'primary education',
                'secondary education', 'tertiary education', 'gender', 'women',
                'girl', 'queer', 'female', 'feminis', 'non-binary',
                'non binary', 'sexes', 'lgbtq', 'patriarchy', 'transgender',
                'two-spirit', 'gender equality', 'violence against women',
                'trafficking', 'forced marriage', 'water', 'sanita',
                'contamination', 'arid', 'drought', 'hygien', 'sewage',
                'water scarcity', 'remediation', 'stormwater management',
                'low impact development', 'green infrastructure',
                'living infrastructure', 'water education',
                'untreated wastewater', 'water harvesting', 'desalination',
                'water efficiency', 'groundwater depletion', 'desertification',
                'water filtration', 'latrines', 'open defecation',
                'hydrological cycle', 'water nexus', 'energy nexus', 'energy',
                'renewabl', 'wind', 'solar', 'geothermal', 'hydroelectric',
                'fuel efficient', 'fuel-efficient', 'carbon capture',
                'emission', 'greenhouse', 'biofuel', 'energy sovereignty',
                'energy security', 'energy education', 'employment',
                'economic growth', 'sustainable development', 'labour', 'labor',
                'worker', 'wage', 'economic empowerment', 'entrepreneur',
                'small-sized enterprise', 'medium-sized enterprise', 'smes',
                'sustainable tourism', 'youth employment', 'green job',
                'economic recovery', 'green growth', 'sustainable growth',
                'infrastructure', 'buildings', 'capital', 'invest', 'internet',
                'globaliz', 'globalis', 'industrialization', 'value chain',
                'affordable credit', 'industrial diversification', 'trade',
                'inequality', 'financial market', 'taxation', 'equit',
                'equalit', 'humanitarian', 'minorit', 'refugee', 'bipoc',
                'of colour', 'of color', 'indigenous', 'reconciliation',
                'truth and reconciliation', 'underserved', 'privileged',
                'affordab', 'equal access', 'marginalized', 'marginalised',
                'impoverished', 'vulnerable population', 'social safety',
                'social security', 'government program', 'disparity', 'income',
                'gini', 'anti-oppressive', 'anti-racist', 'anti-discriminatory',
                'decolonization', 'cities', 'urban', 'resilien', 'rural',
                'sustainable development', 'public transport', 'metro',
                'housing', 'green infrastructure', 'low impact development',
                'climate change adaptation', 'climate change mitigation',
                'green buildings', 'affordable housing', 'walkab', 'transit',
                'civic spaces', 'open spaces', 'accessib',
                'indigenous placemaking', 'indigenous placekeeping', 'consum',
                'production', 'waste', 'natural resource', 'recycl',
                'industrial ecology', 'sustainable design', 'supply chain',
                'outsourc', 'offshor', 'reuse', 'decarboniz', 'decarbonis',
                'carbon tax', 'carbon pricing', 'food waste',
                'public procurement', 'fossil fuel subsidies', 'climate',
                'greenhouse gas', 'global warming', 'weather', 'environmental',
                'planet', 'vegan', 'vegetarian', 'anthropogenic', 'fossil fuel',
                'emissions', 'carbon dioxide', 'co2', 'carbon-neutral',
                'carbon neutral', 'net zero', 'net-zero', 'methane',
                'sea level', 'climate change mitigation',
                'climate change adaptation', 'climate impacts',
                'climate scenarios', 'climate solutions', 'climate justice',
                'global climate models', 'carbon capture',
                'carbon sequestration', 'low carbon', 'resilience',
                'anthropocene', 'climate positive', 'offsets', 'carbon trading',
                'carbon markets', 'unfccc', 'climate finance',
                'loss and damage', 'paris agreement', 'kyoto protocol',
                'convention of the parties', 'environmental racism',
                'environmental justice', 'eco justice', 'sacrifice zone',
                'political ecology', 'toxic waste', 'ecocide',
                'environmentalism of the poor', 'ocean', 'marine', 'pollut',
                'conserv', 'fish', 'natural habitat', 'species', 'animal',
                'biodivers', 'coral', 'maritime', 'ocean literacy', 'ecosystem',
                'overfish', 'fish stocks', 'ocean', 'sustainable use',
                'traditional use', 'forest', 'biodivers', 'ecolog', 'pollut',
                'conserv', 'land use', 'natural habitat', 'species', 'animal',
                'regeneration', 'resilience', 'sustainable and traditional use',
                'land', 'ecological restoration', 'forest conservation',
                'carbon sequestration', 'carbon capture', 'soil', 'erosion',
                'habitat loss', 'endangered species', 'ecosystem',
                'deforestation', 'reforestation', 'wildlife', 'flora and fauna',
                'benefit sharing', 'institut', 'governance', 'peace',
                'social justice', 'injustice', 'criminal justice',
                'human rights', 'democratic rights', 'voter rights',
                'legal system', 'social change', 'corrupt', 'nationalism',
                'democra', 'authoritarian', 'indigenous', 'judic', 'ecojustice',
                'eco justice', 'indigenous rights', 'self-determination',
                'sovereignty', 'violence', 'exploitation', 'trafficking',
                'torture', 'rule of law', 'illicit', 'organized crime', 'bribe',
                'terroris', 'prior and informed consent',
                'access and benefit sharing', 'undrip',
                'united nations declaration on rights of indigenous peoples']


# dict of all SDG keywords and which SDG group they belong to
SDG_groups = {
    'SDG1': ['poverty', 'income distribution', 'wealth distribution', 'socio '
                                                                      'economic',
             'socio-economic', 'socioeconomic', 'homeless', 'low-income',
             'low income', 'affordab', 'disparity', 'welfare', 'social safety',
             'developing country', 'vulnerability', 'precarity', 'precarious',
             'pro-poor'],
    'SDG2': ['agricultur', 'nutrition', 'food security', 'food insecurity',
             'food-secure', 'food system', 'child hunger', 'food justice',
             'food scarcity', 'food sovereignty', 'food culture', 'culinary',
             'malnutrition', 'agro', 'permaculture', 'indigenous crops',
             'regenerative agriculture', 'urban agriculture', 'organic food',
             'biodynamic', 'food literacy', 'food education', 'benefit sharing',
             'access and benefit sharing', 'ABS', 'end hunger', 'food price',
             'zero hunger'],
    'SDG3': ['well being', 'wellbeing', 'well-being',
             'mental health', 'public health', 'global health', 'health care',
             'healthcare', 'health issues', 'mental wellness', 'disabilit',
             'sexual education', 'mindfulness', 'holism', 'illness',
             'health education', 'communicable disease', 'health determinants',
             'vaccine', 'substance abuse', 'maternal mortality',
             'family planning', 'hazardous chemicals', 'pollution',
             'health equity', 'neonatal mortality', 'infant mortality',
             'child health', 'road traffic accidents', 'reproductive health',
             'epidemics', 'universal health coverage'],
    'SDG4': ['equitable', 'pedagogy', 'knowledge', 'worldview', 'learning',
             'knowledges', 'traditional knowledge', 'land-based knowledge',
             'place-based knowledge', 'decolonial', 'anticolonial', 'settler',
             'equitable', 'equity', 'anti-racism', 'racism', 'anti-oppression',
             'oppression,' 'anti-discriminatory', 'early childhood development',
             'peace', 'citizen', 'sustainability teaching',
             'sustainability education', 'universal literacy', 'basic literacy',
             'universal numeracy', 'environmental education',
             'education for sustainable development', 'ecojustice education',
             'eco justice education',
             'place-based education', 'humane education', 'land-based learning',
             'nature-based education', 'climate change education', 'vocational',
             'technical learning', 'free education', 'accessible education',
             'primary education', 'secondary education', 'tertiary education'],
    'SDG5': ['gender', 'women', 'girl', 'queer', 'female', 'feminis',
             'non-binary', 'non binary', 'sexes', 'lgbtq', 'patriarchy',
             'transgender', 'two-spirit', 'gender equality',
             'violence against women', 'trafficking', 'forced marriage'],
    'SDG6': ['water', 'sanita', 'contamination', 'arid', 'drought', 'hygien',
             'sewage', 'water scarcity', 'remediation', 'stormwater management',
             'low impact development', 'green infrastructure',
             'living infrastructure', 'water education', 'untreated wastewater',
             'water harvesting', 'desalination', 'water efficiency',
             'groundwater depletion', 'desertification', 'water filtration',
             'latrines', 'open defecation', 'hydrological cycle',
             'water nexus', 'energy nexus'],
    'SDG7': ['energy', 'renewabl', 'wind', 'solar', 'geothermal',
             'hydroelectric', 'fuel efficient', 'fuel-efficient', 'carbon '
                                                                  'capture',
             'emission', 'greenhouse', 'biofuel', 'energy sovereignty', ''
                                                                        'energy security',
             'energy education'],
    'SDG8': ['employment', 'economic growth', 'sustainable development',
             'labour', 'labor', 'worker', 'wage', 'economic empowerment',
             'entrepreneur', 'small-sized enterprise',
             'medium-sized enterprise', 'smes',
             'sustainable tourism', 'youth employment', 'green job',
             'economic recovery', 'green growth', 'sustainable growth'],
    'SDG9': ['infrastructure', 'buildings', 'capital', 'invest', 'internet',
             'globaliz', 'globalis', 'industrialization', 'value chain',
             'affordable credit', 'industrial diversification'],
    'SDG10': ['trade', 'inequality', 'financial market', 'taxation', 'equit',
              'equalit', 'humanitarian', 'minorit', 'refugee', 'bipoc',
              'of colour', 'of color', 'indigenous', 'reconciliation',
              'truth and reconciliation', 'underserved', 'privileged',
              'affordab', 'equal access', 'marginalized', 'marginalised',
              'impoverished', 'vulnerable population', 'social safety',
              'social security', 'government program', 'disparity', 'income',
              'gini', 'anti-oppressive', 'anti-racist', 'anti-discriminatory',
              'decolonization'],
    'SDG11': ['cities', 'urban', 'resilien', 'rural', 'sustainable '
                                                      'development',
              'public transport', 'metro', 'housing', 'green infrastructure',
              'low impact development', 'climate change adaptation',
              'climate change mitigation', 'green buildings',
              'affordable housing', 'walkab', 'transit', 'civic spaces',
              'open spaces', 'accessib', 'indigenous placemaking',
              'indigenous placekeeping'],
    'SDG12': ['consum', 'production', 'waste', 'natural resource', 'recycl',
              'industrial ecology',
              'sustainable design', 'supply chain', 'outsourc', 'offshor',
              'reuse', 'decarboniz', 'decarbonis', 'carbon tax', 'carbon pricing',
              'food waste', 'public procurement', 'fossil fuel subsidies'],
    'SDG13': ['climate', 'greenhouse gas', 'global warming',
              'weather', 'environmental', 'planet', 'vegan', 'vegetarian',
              'anthropogenic', 'fossil fuel', 'emissions', 'carbon dioxide',
              'co2', 'carbon-neutral', 'carbon neutral', 'net zero', 'net-zero',
              'methane', 'sea level', 'climate change mitigation',
              'climate change adaptation', 'climate impacts',
              'climate scenarios', 'climate solutions', 'climate justice',
              'global climate models', 'carbon capture', 'carbon sequestration',
              'low carbon', 'resilience', 'anthropocene', 'climate positive',
              'offsets', 'carbon trading', 'carbon markets', 'unfccc',
              'climate finance', 'loss and damage', 'paris agreement',
              'kyoto protocol', 'convention of the parties',
              'environmental racism', 'environmental justice', 'eco justice',
              'sacrifice zone', 'political ecology', 'toxic waste', 'ecocide',
              'environmentalism of the poor'],
    'SDG14': ['ocean', 'marine', 'pollut', 'conserv', 'fish', 'natural habitat',
              'species', 'animal', 'biodivers', 'coral', 'maritime',
              'ocean literacy', 'ecosystem', 'overfish', 'fish stocks', 'ocean',
              'sustainable use', 'traditional use'],
    'SDG15': ['forest', 'biodivers', 'ecolog', 'pollut', 'conserv',
              'land use', 'natural habitat', 'species', 'animal',
              'regeneration', 'resilience', 'sustainable and traditional use',
              'land', 'ecological restoration', 'forest conservation',
              'carbon sequestration', 'carbon capture', 'soil', 'erosion',
              'habitat loss', 'endangered species', 'ecosystem',
              'deforestation', 'reforestation', 'wildlife', 'flora and fauna',
              'benefit sharing'],
    'SDG16': ['institut', 'governance', 'peace',
              'social justice', 'injustice', 'criminal justice', 'human rights',
              'democratic rights', 'voter rights', 'legal system',
              'social change', 'corrupt', 'nationalism', 'democra',
              'authoritarian', 'indigenous', 'judic', 'ecojustice',
              'eco justice',
              'indigenous rights', 'self-determination', 'sovereignty',
              'violence', 'exploitation', 'trafficking', 'torture',
              'rule of law', 'illicit', 'organized crime', 'bribe', 'terroris',
              'prior and informed consent', 'access and benefit sharing',
              'undrip',
              'united nations declaration on rights of indigenous peoples']}

write_sheet(sheet_path, wb_name, titles, col_width)
