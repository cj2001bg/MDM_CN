# 存储参数

# 间接物料的所有字段名称
def SPfields():
    sapFields = [
    'Line Number',
    'Material Number',
    'Material Description (English)',
    'Material Description (Chinese)',
    'Industry Sector',
    'Material Type',
    'Plant',
    'Storage Location',
    'Sales Organization',
    'Distribution Channel',
    'Base Unit of Measure',
    'Material Group',
    'Old Material Number',
    'External Material Group',
    'Cross-Plant Material Status',
    'Gross Weight',
    'Weight Unit',
    'Net Weight',
    'Volume',
    'Volume Unit',
    'Size/Dimensions',
    'Material Group: Packaging Materials',
    'Basic Material',
    'Industry Std Desc',
    'Class Type',
    'Class',
    'Country Key',
    'Tax Data - Tax Category',
    'Tax Data - Tax Class 1',
    'Cash Discount Indicator',
    'Material Statistics Group',
    'Account Assignment Group',
    'General Item Category Group',
    'Item Category Group',
    'Material Group 1',
    'Material Group 2',
    'Material Group 3',
    'Material Group 4',
    'Material Group 5',
    'Availability Check',
    'Transportation Group',
    'Loading Group',
    'Delivering Plant',
    'Material Pricing Group',
    'Product Hierarchy',
    'Stock Determination Group',
    'Country of Origin',
    'Order Unit',
    'Var. Oun',
    'Purchasing Group',
    'Purchasing Value Key',
    'Batch Management Requirement Ind',
    'Source List Requirement',
    'MRP Group',
    'Plant-Specific Material Status',
    'ABC Indicator',
    'MRP Type',
    'MRP Controller',
    'Lot Size',
    'Minimum Lot Size',
    'Maximum Lot Size',
    'Maximum Stock Level',
    'Rounding Value for Purc Order Qty',
    'Procurement Type',
    'Special Procurement Type',
    'Backflush',
    'Issue Storage Location',
    'Default Storage Location',
    'In-house Production Time',
    'Planned Delivery Time in Days',
    'Goods Receipt Processing Time in Days',
    'Scheduling Margin Key for Floats',
    'Safety Stock',
    'Period Indicator',
    'Fiscal Year Variant',
    'Planning Strategy Group',
    'Consumption Mode',
    'BWD Consumption Period',
    'Component Scrap (%)',
    'Under Tolerance',
    'Over Tolerance',
    'Unit of Issue',
    'Storage Bin',
    'Storage Condition',
    'CC Phys Inventory Indicator',
    'Minimum Remaining Shelf Life',
    'Total Shelf Life',
    'Period indicator for SLED',
    'Negative Stocks Allowed in Plant',
    'MRP Relevancy for Dependent Requirements',
    'Valuation Class',
    'Valuation Category',
    'Price Control Indicator',
    'Price Unit',
    'Standard Price',
    'Moving Average Price/Periodic Unit Price',
    'Future Price',
    'Future Price - Valid From (YYYYMMDD)',
    'BOM Usage',
    'Alternative BOM',
    'Lot Size for Product Costing',
    'Price Determination',
    'Material Ledger Indicator',
    'With Quantity Structure',
    'Material Origin',
    'Variance Key'
    ]
    return sapFields

# 中文国家名称对应的SAP代码字典
def CountryDict():
    # 间接物料汇总常用国家字典
    countryDict = {
            '美国':'US',
            '英国':'GB',
            '中国':'CN',
            '意大利':'IT',
            '德国':'DE',
            '瑞士':'CH',
            '瑞典':'SE',
            '日本':'JP',
            'China':'CN',
            'CHINA':'CN',
            'Germany':'DE',
            'ITALY':'IT',
            '加拿大':'CA',
            'Netherlands':'NL'
    }
    return countryDict

# POTXT所有字段名称
def PoTxtList():
    poTxtList = [
    # 'Line Number',
    'Material',
    'Chinese Matl Desc',
    'English Matl Desc',
    'Text Language',
    'Purchase Order Text'
    ]
    return poTxtList
# 工厂对应
def PlantDict():
    plantDict = {
        'SH1':'0078',
        'SH2':'0079',
        'SZ':'0023',
        'SH1,SH2':'0078,0079'
    }
    return plantDict

# Z1UMM24中字段类型定义
def FieldSAP():
    fieldSAP = {
                'Material Description (English)':	str,
                'Material Description (Chinese)':	str,
                'Industry Sector':	str,
                'Material Type':	str,
                'Plant':	str,
                'Storage Location':	str,
                'Sales Organization':	str,
                'Distribution Channel':	str,
                'Base Unit of Measure':	str,
                'Material Group':	str,
                'Old Material Number':	str,
                'External Material Group':	str,
                'Cross-Plant Material Status':	str,
                'Gross Weight':	str,
                'Weight Unit':	str,
                'Net Weight':	str,
                'Volume':	str,
                'Volume Unit':	str,
                'Size/Dimensions':	str,
                'Material Group: Packaging Materials':	str,
                'Basic Material':	str,
                'Industry Std Desc':	str,
                'Class Type':	str,
                'Class':	str,
                'Country Key':	str,
                'Tax Data - Tax Category':	str,
                'Tax Data - Tax Class 1':	str,
                'Cash Discount Indicator':	str,
                'Material Statistics Group':	str,
                'Account Assignment Group':	str,
                'General Item Category Group':	str,
                'Item Category Group':	str,
                'Material Group 1':	str,
                'Material Group 2':	str,
                'Material Group 3':	str,
                'Material Group 4':	str,
                'Material Group 5':	str,
                'Availability Check':	str,
                'Transportation Group':	str,
                'Loading Group':	str,
                'Delivering Plant':	str,
                'Material Pricing Group':	str,
                'Product Hierarchy':	str,
                'Stock Determination Group':	str,
                'Country of Origin':	str,
                'Order Unit':	str,
                'Var. Oun':	str,
                'Purchasing Group':	str,
                'Purchasing Value Key':	str,
                'Batch Management Requirement Ind':	str,
                'Source List Requirement':	str,
                'MRP Group':	str,
                'Plant-Specific Material Status':	str,
                'ABC Indicator':	str,
                'MRP Type':	str,
                'MRP Controller':	str,
                'Lot Size':	str,
                'Minimum Lot Size':	str,
                'Maximum Lot Size':	str,
                'Maximum Stock Level':	str,
                'Rounding Value for Purc Order Qty':	str,
                'Procurement Type':	str,
                'Special Procurement Type':	str,
                'Backflush':	str,
                'Issue Storage Location':	str,
                'Default Storage Location':	str,
                'In-house Production Time':	str,
                'Planned Delivery Time in Days':	str,
                'Goods Receipt Processing Time in Days':	str,
                'Scheduling Margin Key for Floats':	str,
                'Safety Stock':	str,
                'Period Indicator':	str,
                'Fiscal Year Variant':	str,
                'Planning Strategy Group':	str,
                'Consumption mode':	str,
                'BWD Consumption Period':	str,
                'Component Scrap (%)':	str,
                'Under Tolerance':	str,
                'Over Tolerance':	str,
                'Unit of Issue':	str,
                'Storage Bin':	str,
                'Storage Condition':	str,
                'CC Phys Inventory Indicator':	str,
                'Period indicator for SLED':	str,
                'Negative Stocks Allowed in Plant':	str,
                'MRP Relevancy for Dependent Requirements':	str,
                'Valuation Class':	str,
                'Valuation Category':	str,
                'Price Control Indicator':	str,
                'Price Unit':	str,
                'Standard Price':	str,
                'Moving Average Price/Periodic Unit Price':	str,
                'Future Price':	str,
                'Future Price - Valid From (YYYYMMDD)':	str,
                'BOM Usage':	str,
                'Alternative BOM':	str,
                'Lot Size for Product Costing':	str,
                'Price Determination':	str,
                'Material Ledger Indicator':	str,
                'With Quantity Structure':	str,
                'Material Origin':	str,
                'Variance Key':	str
                    }
    return fieldSAP

# Source List 字段类型定义
def FieldSL():
    fieldSL = {
                'Plant':	str,
                'Record No.':	str,
                'Valid from':	str,
                'Valid to':	str,
                'Vendor':	str,
                'Purch Org':	str,
                'Fixed Vendor':	str,
                'Block Source of Supp':	str,
                'MRP Indicator':	str
                    }

    return fieldSL
    
# Acc字段类型定义
def FieldDictACC():
    fieldDict2 = {
                'English Matl Desc':	str,
                'Chinese Matl Desc':	str,
                'Plant':	str,
                'Valuation Key':	str,
                'Valuation Type':	str,
                'Valuation Category':	str,
                'Price Determination':	str,
                'Material Ledger Indicator':	str,
                'Valuation Class':	str,
                'VC:Sales Order Stk':	str,
                'Price Control Indicator':	str,
                'Price Unit':	str,
                'Standard Price':	str,
                'Moving Average Price/Periodic Unit Price':	str,
                'Future Price':	str,
                'Future Price - Valid From (YYYYMMDD)':	str,
                'With Quantity Structure':	str,
                'Material Origin':	str,
                'Variance Key':	str,
                'Alternative BOM':	str,
                'BOM Usage':	str,
                'Lot Size for Product Costing':	str
                }
    return fieldDict2

# Acc上传模板字段列表
def FieldACCList():
    fieldACCList = [
                    'Material Code',
                    'English Matl Desc',
                    'Chinese Matl Desc',
                    'Plant',
                    'Valuation Key',
                    'Valuation Type',
                    'Valuation Category',
                    'Price Determination',
                    'Material Ledger Indicator',
                    'Valuation Class',
                    'VC:Sales Order Stk',
                    'Price Control Indicator',
                    'Price Unit',
                    'Standard Price',
                    'Moving Average Price/Periodic Unit Price',
                    'Future Price',
                    'Future Price - Valid From (YYYYMMDD)',
                    'With Quantity Structure',
                    'Material Origin',
                    'Variance Key',
                    'Alternative BOM',
                    'BOM Usage',
                    'Lot Size for Product Costing'
    ]
    return fieldACCList

# UOM上传模板字段列表
def FieldUOMList():
    fieldUOMList = [
                    'Material Code',
                    'English Matl Desc',
                    'Chinese Matl Desc',
                    'Conversion Value (X)',
                    'Alternative UOM',
                    'Conversion Value (Y)',
                    'Base UOM',
                    'Delete?',
                    'EAN/UPC',
                    'EAN CAT'
                    ]
    return fieldUOMList

# UOM字段类型定义
def FieldUOM():
    fieldUOM = {
            'English Matl Desc':	str,
            'Chinese Matl Desc':	str,
            'Conversion Value (X)':	str,
            'Alternative UOM':	str,
            'Conversion Value (Y)':	str,
            'Base UOM':	str,
            'Delete?':	str
                }
    return fieldUOM

# Z1UMM24中所有字段列表
def FieldSAPALL():
    fieldSAPList = [
                    'Material Number',
                    'Material Description (English)',
                    'Material Description (Chinese)',
                    'Industry Sector',
                    'Material Type',
                    'Plant',
                    'Storage Location',
                    'Sales Organization',
                    'Distribution Channel',
                    'Base Unit of Measure',
                    'Material Group',
                    'Old Material Number',
                    'External Material Group',
                    'Cross-Plant Material Status',
                    'Gross Weight',
                    'Weight Unit',
                    'Net Weight',
                    'Volume',
                    'Volume Unit',
                    'Size/Dimensions',
                    'Material Group: Packaging Materials',
                    'Basic Material',
                    'Industry Std Desc',
                    'Class Type',
                    'Class',
                    'Country Key',
                    'Tax Data - Tax Category',
                    'Tax Data - Tax Class 1',
                    'Cash Discount Indicator',
                    'Material Statistics Group',
                    'Account Assignment Group',
                    'General Item Category Group',
                    'Item Category Group',
                    'Material Group 1',
                    'Material Group 2',
                    'Material Group 3',
                    'Material Group 4',
                    'Material Group 5',
                    'Availability Check',
                    'Transportation Group',
                    'Loading Group',
                    'Delivering Plant',
                    'Material Pricing Group',
                    'Product Hierarchy',
                    'Stock Determination Group',
                    'Country of Origin',
                    'Order Unit',
                    'Var. Oun',
                    'Purchasing Group',
                    'Purchasing Value Key',
                    'Batch Management Requirement Ind',
                    'Source List Requirement',
                    'MRP Group',
                    'Plant-Specific Material Status',
                    'ABC Indicator',
                    'MRP Type',
                    'MRP Controller',
                    'Lot Size',
                    'Minimum Lot Size',
                    'Maximum Lot Size',
                    'Maximum Stock Level',
                    'Rounding Value for Purc Order Qty',
                    'Procurement Type',
                    'Special Procurement Type',
                    'Backflush',
                    'Issue Storage Location',
                    'Default Storage Location',
                    'In-house Production Time',
                    'Planned Delivery Time in Days',
                    'Goods Receipt Processing Time in Days',
                    'Scheduling Margin Key for Floats',
                    'Safety Stock',
                    'Period Indicator',
                    'Fiscal Year Variant',
                    'Planning Strategy Group',
                    'Consumption Mode',
                    'BWD Consumption Period',
                    'Component Scrap (%)',
                    'Under Tolerance',
                    'Over Tolerance',
                    'Unit of Issue',
                    'Storage Bin',
                    'Storage Condition',
                    'CC Phys Inventory Indicator',
                    'Minimum Remaining Shelf Life',
                    'Total Shelf Life',
                    'Period indicator for SLED',
                    'Negative Stocks Allowed in Plant',
                    'MRP Relevancy for Dependent Requirements',
                    'Valuation Class',
                    'Valuation Category',
                    'Price Control Indicator',
                    'Price Unit',
                    'Standard Price',
                    'Moving Average Price/Periodic Unit Price',
                    'Future Price',
                    'Future Price - Valid From (YYYYMMDD)',
                    'BOM Usage',
                    'Alternative BOM',
                    'Lot Size for Product Costing',
                    'Price Determination',
                    'Material Ledger Indicator',
                    'With Quantity Structure',
                    'Material Origin',
                    'Variance Key'
    ]
    return fieldSAPList

# Info.字段类型定义
def FieldInfo():
    fieldInfo = {
                'Material Description':	str,
                'Purch. Org.':	str,
                'Plant':	str,
                'Incoterms':	str,
                'Incoterms Text':	str,
                'Currency':	str,
                'Order Unit':	str,
                'Tax Code':	str,
                'GR-based IV':	str,
                'Valid On':	str,
                'Valid To':	str,
                'Condition Type':	str,
                'Rate Currency':	str,
                'UoM':	str
                    }
    return fieldInfo

# FG HANA上传模板字段列表SAP
def FGHanaFields():
    fgHanaFields = [
            'Material',	'Class',	'Category',	'Net. Decl. Weight',	'Weight Per Piece',	'Num Piece per Pack',	'Segment',
            'Subsegment',	'Brand',	'Sub Brand',	'Flavour group',	'Flavour',	'Sugar Content',	'Manufacturing Type',
            'Production Base Type',	'Production Technology Group',	'Promo',	'Launch Date',	'Single piece wrapper',
        	'Primary pack type',	'Secondary pack type',	'Additional pack remark',	'Fact Prod',	'GL_FACT_PACKING',
            'Coating',	'Filling'
                        ]
    return fgHanaFields

# FG SAP-HANA对应
def FGHanaFields2():
    fgHanaFields2 = {
                    'Material':	'Material',
                    'Category':	'Category',
                    'Net. Decl. Weight':	'Net Decl. Weight',
                    'Weight Per Piece':	'Weight per Piece',
                    'Num Piece per Pack':	'Num Piece per Pack',
                    'Segment':	'Segment',
                    'Subsegment':	'Sub Segment',
                    'Brand':	'Brand',
                    'Sub Brand':	'Sub Brand / Range',
                    'Flavour group':	'Flavour Group',
                    'Flavour':	'Flavour',
                    # 'Sugar Content':	'Sugar Content',
                    # 'Production Base Type':	'Prod Base Type',
                    # 'Production Technology Group':	'Prod Tech Group',
                    'Single piece wrapper':	'Single Piece Wrapper',
                    'Primary pack type':	'Primary Pack Type',
                    'Secondary pack type':	'Secondary Pack Type',
                    'Additional pack remark':	'Additional Packaging Remark',
                    'Fact Prod':	'Factory Production',
                    'GL_FACT_PACKING':	'Factory Packing'
                    # 'Coating':	'Coating',
                    # 'Filling':	'Filling'
                    }
    return fgHanaFields2

# FG PAF-HANA字段对应for copy
def Paf_hana1():
    paf_hana = {
                'Material':'New SKU Code',
                'Category':'Category',
                'Segment':'Segment',
                'Sub Segment':'Sub Segment',
                'Brand':'Brand',
                'Single Piece Wrapper':'Single Piece Wrapper',
                'Primary Pack Type':'Primary Packaging',
                'Secondary Pack Type':'Second Packaging',
                'Additional Packaging Remark':'Additional Packaging Remark',
                'Flavour Group':'Flavor Group',
                'Flavour':'Flavor',
                'Sub Brand / Range':'Sub Brand /Range/ Product Name'
                }
    return paf_hana

# FG HANA Report字段列表
def Hana_list():
    hana_list = [
                'Status',
                'Source System',
                'Material Type',
                'Material',
                'Description Short',
                'Description Long',
                'EAN Code',
                'Category',
                'Segment',
                'Sub Segment',
                'Brand',
                'Sub Brand / Range',
                'NPD',
                'Factory Production',
                'Factory Packing',
                'Net Decl. Weight',
                'Calculation Weight',
                'Weight per Piece',
                'Pack Size Grms Nielsen',
                'Num Piece per Pack',
                'Manuf. Type',
                'Coating',
                'Filling',
                'Sugar Content',
                'Prod Base Type',
                'Base Type Prop.',
                'Prod Tech Group',
                'Prod Tech Type',
                'First Selling OC',
                'First Invoice Date - Date',
                'Single Piece Wrapper',
                'Primary Pack Type',
                'Secondary Pack Type',
                'Additional Packaging Remark',
                'Flavour Group',
                'Flavour',
                'Count Material',
                'DQ Status'
                ]
    return hana_list

# R/PM SAP_class字段列表
def rpm_class_list():
    rpm_class_list = [
        'Material',
        'Class',
        'Material type',
        'Category 1',
        'Category 2',
        'Category 3',
        'Material structure',
        'Material structure 2'
    ]
    return rpm_class_list
