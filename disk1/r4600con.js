//
// Copyright (c) 2016-2017 Ricoh Co., Ltd. All rights reserved.
//
// Abstract:
//    RICOH V4 Printer Driver
//

var psfPrefix = "psf";
var pskPrefix = "psk";
var userpskPrefix = "ns0000";

var pskNs = "http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords";
var psfNs = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework";
var ricohNs = "http://schemas.ricoh.com/2007/printing/keywords";

var NsMap = {
    "psk":"http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords",
    "ns0000":"http://schemas.ricoh.com/2007/printing/keywords"
};

var FRAMEWORK_URI = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework";
var KEYWORDS_URI = "http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords";

var SCHEMA_INST_URI = "http://www.w3.org/2001/XMLSchema-instance";
var SCHEMA_DEF_URI = "http://www.w3.org/2001/XMLSchema";


var SCHEMA_XSI = "xsi";
var SCHEMA_XS = "xsd";

var SCHEMA_TYPE = "type";
var SCHEMA_INTEGER = "integer";
var SCHEMA_DECIMAL = "decimal";
var SCHEMA_STRING = "string";
var SCHEMA_QNAME = "QName";
var SCHEMA_CONDITIONAL = "Conditional";

var FEATURE_ELEMENT_NAME = "Feature";
var OPTION_ELEMENT_NAME = "Option";
var PARAM_INIT_ELEMENT_NAME = "ParameterInit";
var PARAM_REF_ELEMENT_NAME = "ParameterRef";
var PARAM_DEF_ELEMENT_NAME = "ParameterDef";
var SCORED_PROP_ELEMENT_NAME = "ScoredProperty";
var PROPERTY_ELEMENT_NAME = "Property";
var VALUE_ELEMENT_NAME = "Value";
var NAME_ATTRIBUTE_NAME = "name";

var PICKONE_VALUE_NAME = "PickOne";
var SELECTIONTYPE_VALUE_NAME = "SelectionType";
var DATATYPE_VALUE_NAME = "DataType";
var DEFAULTVAL_VALUE_NAME = "DefaultValue";
var MAX_VALUE_NAME = "MaxValue";
var MIN_VALUE_NAME = "MinValue";
var MAX_LENGTH_NAME = "MaxLength";
var MIN_LENGTH_NAME = "MinLength";
var MULTIPLE_VALUE_NAME = "Multiple";
var MANDATORY_VALUE_NAME = "Mandatory";
var UNITTYPE_VALUE_NAME = "UnitType";
var DISPLAYNAME_VALUE_NAME = "DisplayName";

var PRINTTICKET_NAME = "PrintTicket";
var PRINTCAPABILITIES_NAME = "PrintCapabilities";

var PROPERTY_KEY_ENABLE_CHANGEDEFAULT = "JS_Customize_Enabled";
var PROPERTY_KEY_MODE_MODELCHANGED    = "JS_Model_Changed";

var EVENT_DEVMODETOPRINTTICKET = "1";
var EVENT_PRINTTICKETTODEVMODE = "2";
var EVENT_PRINTCAPABILITIES    = "3";

var FEATURE_PAGE_ORIENTATION   = "psk:PageOrientation";
var OPTION_LANDSCAPE = "psk:Landscape";
var OPTION_REVERSELANDSCAPE = "psk:ReverseLandscape";

var global = {};
global.eventtype = null;
global.language  = "unknown";
global.isDefaultRequest  = false;

var AUTO_CREATE_FEATURE_LIST = [
    "psk:PageOrientation",
    "psk:PageMediaSize",
    "psk:PageOutputColor",
    "psk:JobNUpAllDocumentsContiguously",
    "psk:JobStapleAllDocuments",
    "psk:JobHolePunch",
    "psk:PageMediaType",
    "psk:JobDuplexAllDocumentsContiguously",
    "psk:PageInputBin",
    "psk:DocumentCollate",
    "psk:JobPageOrder",
    "psk:JobOutputBin",
    "ns0000:JobBooklet",
    "ns0000:JobGraphicsMode",
    "ns0000:JobPrintableArea",
    "ns0000:JobImageDirection",
    "ns0000:JobEdgeToEdgePrint"
];

var AUTO_CREATE_PARAMETERDEF_LIST = [
    "psk:JobCopiesAllDocuments",
    "psk:PageMediaSizeMediaSizeWidth",
    "psk:PageMediaSizeMediaSizeHeight"
];

var AUTO_CREATE_SCOREDPROPERTY_LIST = [
    "psk:PagesPerSheet"
];

var DUPLICATED_FEATURENAME_LIST = [
    "ns0000:ProfileName",
    "ns0000:Border",
];

var RELATED_FEATURE_LIST = {
    "ns0000:WatermarkItem":
    {
        "name":"ns0000:PageWatermarkSettings",
        "default":""
    },
    "ns0000:OverlayWatermarkItem":
    {
        "name":"ns0000:PageOverlayWatermarkSettings",
        "default":""
    }
//    "ns0000:PageColorProfileDistribution":
//    {
//        "name":"psk:PageColorManagement",
//        "default":"psk:None",
//        "pickone":["psk:None"]
//    }
};

var CHANGEABLE_PICKONE_FEATURE_LIST = [
    "ns0000:WatermarkItem",
    "ns0000:OverlayWatermarkItem",
    "ns0000:OverlayWatermarkItemAll",
    "ns0000:OverlayWatermarkItem1stPage",
    "ns0000:OverlayWatermarkItem2ndPage",
    "ns0000:OverlayWatermarkItem3rdPage",
    "ns0000:ImageWatermarkItem",
    "ns0000:PageColorProfile",
    "ns0000:PageColorProfileForText",
    "ns0000:PageColorProfileForGraphics",
    "ns0000:PageColorProfileForPhoto"
];

var USERFORM_SUPPORT_TABLE = {
    "ns0000:PerfectBindCutAdjustmentByPaperSize" :
    {
        "width": "ns0000:CutSizeWidth", "height": "ns0000:CutSizeHeight"
    },
    "ns0000:PageStandardPaperSize" :
    {
        "width": "ns0000:StandardMediaSizeWidth", "height": "ns0000:StandardMediaSizeHeight"
    },
    "ns0000:InsertSeparationSheetPaperSize" :
    {
        "width": "ns0000:InsertSeparationSheetMediaSizeWidth", "height": "ns0000:InsertSeparationSheetMediaSizeHeight"
    },
    "ns0000:GlueBindPaperSize" :
    {
        "width": "ns0000:GlueBindMediaSizeWidth", "height": "ns0000:GlueBindMediaSizeHeight"
    },
    "ns0000:InsertJobSeparationSheetPaperSize" :
    {
        "width": "ns0000:InsertJobSeparationSheetMediaSizeWidth", "height": "ns0000:InsertJobSeparationSheetMediaSizeHeight"
    },
    "ns0000:SlipSheetPaperSize" :
    {
        "width": "ns0000:SlipSheetMediaSizeWidth", "height": "ns0000:SlipSheetMediaSizeHeight"
    },
    "ns0000:BookletCoverSheetSettingsPaperSize" :
    {
        "width": "ns0000:BookletCoverSheetMediaSizeWidth", "height": "ns0000:BookletCoverSheetMediaSizeHeight"
    }
};

var CHANGFEATURE_TABLE_LIST = [
    "ns0000:MediaSizeWidth",
    "ns0000:MediaSizeHeight"
];

var PSKJOBDUPLEX_CONTIGUOUSLY_MAP = {
    'ns0000:Off': 'psk:OneSided',
    'ns0000:LongEdgeLeftOrTop': 'psk:TwoSidedLongEdge',
    'ns0000:LongEdgeRightOrTop': 'psk:TwoSidedLongEdge',
    'ns0000:ShortEdgeTopOrLeft': 'psk:TwoSidedShortEdge',
    'ns0000:ShortEdgeTopOrRight': 'psk:TwoSidedShortEdge'
};

var RPSKJOBDUPLEX_DIRECTION_MAP = {
    'psk:OneSided':
        {
            'ns0000:Off': 'ns0000:Off' ,
            'ns0000:LongEdgeLeftOrTop': 'ns0000:Off' ,
            'ns0000:LongEdgeRightOrTop': 'ns0000:Off' ,
            'ns0000:ShortEdgeTopOrLeft': 'ns0000:Off' ,
            'ns0000:ShortEdgeTopOrRight': 'ns0000:Off'
        },
    'psk:TwoSidedLongEdge':
        {
            'ns0000:Off': 'ns0000:LongEdgeLeftOrTop' ,
            'ns0000:LongEdgeLeftOrTop': 'ns0000:LongEdgeLeftOrTop' ,
            'ns0000:LongEdgeRightOrTop': 'ns0000:LongEdgeRightOrTop' ,
            'ns0000:ShortEdgeTopOrLeft': 'ns0000:LongEdgeLeftOrTop' ,
            'ns0000:ShortEdgeTopOrRight': 'ns0000:LongEdgeRightOrTop'
        },
    'psk:TwoSidedShortEdge':
        {
            'ns0000:Off': 'ns0000:ShortEdgeTopOrLeft',
            'ns0000:LongEdgeLeftOrTop': 'ns0000:ShortEdgeTopOrLeft',
            'ns0000:LongEdgeRightOrTop': 'ns0000:ShortEdgeTopOrRight',
            'ns0000:ShortEdgeTopOrLeft': 'ns0000:ShortEdgeTopOrLeft',
            'ns0000:ShortEdgeTopOrRight': 'ns0000:ShortEdgeTopOrRight'
        }
}

var RESOLUTION_MAP = {
    'ns0000:_200x100':
        {
            'ResolutionX': '200',
            'ResolutionY': '200'
        },
    'ns0000:_200dpi':
        {
            'ResolutionX': '200',
            'ResolutionY': '200'
        },
    'ns0000:_300dpi':
        {
            'ResolutionX': '300',
            'ResolutionY': '300'
        },
    'ns0000:_400to200dpi':
        {
            'ResolutionX': '400',
            'ResolutionY': '400'
        },
    'ns0000:_400dpi':
        {
            'ResolutionX': '400',
            'ResolutionY': '400'
        },
    'ns0000:_300to600dpi':
        {
            'ResolutionX': '300',
            'ResolutionY': '300'
        },
    'ns0000:_300x300to38400x600':
        {
            'ResolutionX': '300',
            'ResolutionY': '300'
        },
    'ns0000:_600dpi':
        {
            'ResolutionX': '600',
            'ResolutionY': '600'
        },
    'ns0000:_1200x600':
        {
            'ResolutionX': '600',
            'ResolutionY': '600'
        },
    'ns0000:_600x1200':
        {
            'ResolutionX': '600',
            'ResolutionY': '600'
        },
    'ns0000:_2400x600':
        {
            'ResolutionX': '600',
            'ResolutionY': '600'
        },
    'ns0000:_9600x600':
        {
            'ResolutionX': '600',
            'ResolutionY': '600'
        },
    'ns0000:__38400x600':
        {
            'ResolutionX': '600',
            'ResolutionY': '600'
        },
    'ns0000:_1200dpi':
        {
            'ResolutionX': '1200',
            'ResolutionY': '1200'
        },
    'ns0000:_4800x1200':
        {
            'ResolutionX': '1200',
            'ResolutionY': '1200'
        }
}

var FAX_RESOLUTION_MAP = {
  'ns0000:_200x100': 'availableresolution_200x100',
  'ns0000:_200dpi': 'availableresolution_200',
  'ns0000:_400to200dpi': 'availableresolution_200',
  'ns0000:_400dpi': 'availableresolution_400',
  'ns0000:_600dpi': 'availableresolution_600'
}
var FAX_DEFAULT_AVAILABLE_RESOLUTION_MAP = {
  'availableresolution_200x100': 'off',
  'availableresolution_200': 'on',
  'availableresolution_200x400': 'on',
  'availableresolution_400': 'off',
  'availableresolution_600': 'off'
}
var DEFAULT_FAX_RESOLUTION = 'ns0000:_200dpi';
var DEFAULT_PRIMAX_RESOLUTION = 'ns0000:_600dpi';

var JOBEXTRAUISETTINGS_GENERIC = "'ns0000:JobExtraUISettings':{'default':''}"
var JOBEXTRAUISETTINGS = "\"ns0000:JobExtraUISettings\":{\"default\":\"\"}"
var JOBENFORCEUSERCODE_GENERIC = "'ns0000:JobRicEnforceUserCodeDef':{'default':''}"
var JOBENFORCEUSERCODE = "\"ns0000:JobRicEnforceUserCodeDef\":{\"default\":\"\"}"

var PRINT_QUALITY_ADJUSTMENT   = "ns0000:JobPrintQualityAdjustment";
var PRINT_QUALITY_PRESET       = "ns0000:JobPrintQualityPreset";
var COLOR_PROFILE_SETTING = "ns0000:PageColorProfileSetting";
var COLOR_PROFILE         = "ns0000:PageColorProfile";
var DITHER_SETTING        = "ns0000:JobDitherSetting";
var DITHER                = "ns0000:JobDither";
var REPLACE_MAP_COLOR_PROFILE = {"ns0000:Manual":"ns0000:SetEachObject"};
var REPLACE_MAP_DITHER        = {"ns0000:Manual":"ns0000:SetEachObject"};
var REPLACE_MAP_PRINT_QUALITY_PRESET = {"ns0000:Manual":"ns0000:Advanced","ns0000:Auto":"ns0000:Standard"};

var COLOR_PROFILE_DISTRIBUTION = "ns0000:PageColorProfileDistribution";

var KEY_TAG = "tag";
var KEY_ATTRIBUTE ="attribute";
var KEY_CHILDREN = "children";
var KEY_CHILDFEATURE = "childfeature";
var KEY_XPATH = "xpath";
var KEY_VALUE = "value";
var KEY_DEFAULT = "default";
var KEY_NAME = "name";
var KEY_TYPE = "type";
var KEY_PICKONE = "pickone";
var KEY_BLANK = "blank";
var KEY_MIN = "min";
var KEY_MAX = "max";

var TAG_FEATURE = "psf:Feature";
var TAG_PARAMETERDEF = "psf:ParameterDef";
var TAG_FUNCTION = "function";
var TYPE_STRING = "string";
var TYPE_INTEGER = "integer";

var PTPSD_KEY_PTNAME    = "pt_name";
var PTPSD_KEY_PSDNAME   = "psd_name";
var PTPSD_KEY_ATTRIBUTE = "attribute";
var PTPSD_KEY_ATTRIBUTEMULTI = "attribute_multi";
var PTPSD_KEY_PICKONE   = "pickone";

var RCF_ITEM_KEY_NAME     = "name";
var RCF_ITEM_KEY_DEFAULT  = "default";
var RCF_ITEM_KEY_FIXVALUE = "fixvalue";
var RCF_ITEM_KEY_PICKONE  = "pickone";
var RCF_ITEM_KEY_MAX      = "max";
var RCF_ITEM_KEY_VALUE    = "value";
var RCF_ITEM_KEY_ITEMLIST = "items";

var GENERIC_MODEL_NAME = "Generic Model";
/////////////////////////////////////////////
//
//
// Public I/F
//
/////////////////////////////////////////////

////
//// @param printTicket      Print ticket to be validated
//// @param scriptContext    Script context object
////
//// @return Integer value indicating validation status
//// @retval 1 Valid print ticket
//// @retval 2 Print ticket was modified to make it valid
//// @retval 0 Invalid print ticket (not returned by this example)
function validatePrintTicket(printTicket, scriptContext) {
    debugger;

    return 1;
}

////
//// @param printTicket      Print ticket which may constrain the capabilities
//// @param scriptContext    Script context object
//// @param printCapablities Print capabilities object to be expanded
////
//// @return None.
////

//// </summary>
//// <param name="printTicket" type="IPrintSchemaTicket" mayBeNull="true">
////     If not 'null', the print ticket's settings are used to customize the print capabilities.
//// </param>
//// <param name="scriptContext" type="IPrinterScriptContext">
////     Script context object.
//// </param>
//// <param name="printCapabilities" type="IPrintSchemaCapabilities">
////     Print capabilities object to be customized.
//// </param>
function completePrintCapabilities(printTicket, scriptContext, printCapabilities) {

    debugger;
    if (!printCapabilities) {
        return;
    }

    try {
        global.eventtype = EVENT_PRINTCAPABILITIES;

        setCurrentLanguage( printCapabilities );

        CreateAllPrintCapabilities(printTicket, scriptContext, printCapabilities);

        rewriteDisplayName( printCapabilities );

        validatePrintCapabilities( printCapabilities );
    }
    catch (e) {

        var result = 0;
        result = e.number;
        Debug.write(e.message + "completePrintCapabilities() Error!");        
    }

    return;
}

////
//// @param printTicket      Print ticket to be converted to DevMode
//// @param scriptContext    Script context object
//// @param devMode          The DevMode property bag (Persistent storage for properties/data)
////
//// @return None.
////
function convertPrintTicketToDevMode(printTicket, scriptContext, devModeProperties) {

    debugger;
    try
    {
        global.eventtype = EVENT_PRINTTICKETTODEVMODE;
        var property = createPropertyObject(scriptContext, devModeProperties, printTicket);

        convertAllFeaturesPrintTicketToDevmode(printTicket, scriptContext, devModeProperties, property);
        
        validatePrintTicketToDevmode(printTicket, scriptContext);
    }
    catch (e) {

        var result = 0;
        result = e.number;
        Debug.write(e.message + "convertPrintTicketToDevMode() Error!");        
    }
}

////
//// @param devMode          The DevMode property bag (Persistent storage for properties/data)
//// @param scriptContext    Script context object
//// @param printTicket      The output print ticket
////
//// @return None.
////
function convertDevModeToPrintTicket(devModeProperties, scriptContext, printTicket) {    

    debugger;
    try {
        global.eventtype = EVENT_DEVMODETOPRINTTICKET;
        global.isDefaultRequest = isDefaultRequestEvent(scriptContext);

        createAllPrintTicket(devModeProperties, scriptContext, printTicket);

        validateDevmodeToPrintTicket(printTicket, scriptContext);

    }
    catch (e) {
        var result = 0;
        result = e.number;
        Debug.write(e.message + "convertDevModeToPrintTicket() Error!");
    }
}



/////////////////////////////////////////////
//
//
// Private I/F
//
/////////////////////////////////////////////

function getPrintTicketTreeTable( pcTable, property, scriptContext, userformlist )
{
    var tree = [];
    try
    {
        for(var i=0; i<pcTable.length; i++)
        {
            var element = null;
            if( TAG_FEATURE == pcTable[i][KEY_TAG] )
            {
                element = createFeatureElement( pcTable[i], property, null, userformlist );
                if( element !== null )
                {
                    for( var j=0; j<element.length; j++)
                    {
                        tree.push(element[j]);
                    }
                }
            }
            else if( TAG_PARAMETERDEF === pcTable[i][KEY_TAG] )
            {
                element = createParameterDefElement( pcTable[i], property );
                if( element !== null )
                {
                    tree.push(element);
                }
            }
            else if( TAG_FUNCTION === pcTable[i][KEY_TAG] )
            {
                var func = (new Function('arg1', 'arg2', "return " + pcTable[i][KEY_ATTRIBUTE] + "(arg1, arg2)"));
                element = func( scriptContext, property );
                if( element !== null )
                {
                    tree.push(element);
                }
            }
        }
    }
    catch(e) {
        tree = null;
    }

    return tree;
}

function createFeatureElement( pcTable, property, parentFeatureName, userformlist )
{
    var featureElements = [];
    try
    {
        var feature = null;
        var options = null;
        var childFeature = null;
        var userformOptions = null;

        if( Object.prototype.toString.call(pcTable) === '[object Object]' )
        {
            if( TAG_FEATURE == pcTable[KEY_TAG] &&
                (  property.isSupported( pcTable[KEY_ATTRIBUTE] ) ||
                   ( KEY_PICKONE in pcTable && "blank" in pcTable[KEY_PICKONE] ) ||
                   KEY_CHILDFEATURE in pcTable ) )
            {
                feature = new FeatureElement( property, pcTable[KEY_ATTRIBUTE], parentFeatureName );

                options = createOptionElement( pcTable[KEY_PICKONE], property, feature.getName() );
                if( options !== null )
                {
                    for(var i=0; i<options.length; i++ )
                    {
                        feature.appendChild( options[i] );
                    }
                }

                childFeature  = createFeatureElement( pcTable[KEY_CHILDFEATURE], property, feature.getName(), userformlist );
                if( childFeature )
                {
                    for(var j=0; j<childFeature.length; j++)
                    {
                        feature.appendFeature( childFeature[j] );
                    }
                }
                if( (options !== null && options.length > 0) ||
                    (childFeature !== null && childFeature.length > 0 ) )
                {
                    featureElements.push( feature );
                }
            }
        }
        else if( Object.prototype.toString.call(pcTable) === '[object Array]' )
        {
            for(var i=0; i<pcTable.length; i++)
            {
                if( TAG_FEATURE === pcTable[i][KEY_TAG] &&
                    ( property.isSupported( pcTable[i][KEY_ATTRIBUTE] ) || ( KEY_PICKONE in pcTable[i] && KEY_BLANK in pcTable[i][KEY_PICKONE] ) ) )
                {
                    feature = new FeatureElement( property, pcTable[i][KEY_ATTRIBUTE], parentFeatureName );

                    options = createOptionElement( pcTable[i][KEY_PICKONE], property, feature.getName() );
                    if( options !== null )
                    {
                        for(var j=0; j<options.length; j++ )
                        {
                            feature.appendChild( options[j] );
                        }
                    }

                    if( feature.getName() in USERFORM_SUPPORT_TABLE )
                    {
                        userformOptions = createUserFormOptionElement( property, feature.getName(), userformlist );
                        if( userformOptions !== null )
                        {
                            for(var j=0; j<userformOptions.length; j++ )
                            {
                                feature.appendChild( userformOptions[j] );
                            }
                        }
                    }

                    childFeature  = createFeatureElement( pcTable[i][KEY_CHILDFEATURE], property, feature.getName(), userformlist );
                    if( childFeature && feature )
                    {
                        for(var j=0; j<childFeature.length; j++)
                        {
                            feature.appendFeature( childFeature[j] );
                        }
                    }

                    if( (options !== null && options.length > 0) ||
                        (childFeature !== null && childFeature.length > 0 ) )
                    {
                        featureElements.push( feature );
                    }
                }
            }
        }
    }
    catch(e) {
        featureElements = null;
    }
        
    return featureElements;
}


function createOptionElement( optionTable, property, featureName )
{
    var options = [];
    try
    {
        var pickoneList = property.getPickoneList(featureName);

        var option = null;
        var scoredElements = null;
        var optionPickone = null;
        
        if( pickoneList === null )
        {
            if( KEY_BLANK in optionTable )
            {
                option = new OptionElement( "", property );

                if( option !== null )
                {
                    scoredElements = createScoredPropertyElement( optionTable[KEY_BLANK], property, featureName );
                    if( scoredElements.length > 0 )
                    {
                        if( scoredElements.length >= 2 &&
                            scoredElements[0].getName() == scoredElements[1].getName() )
                        {
                            for(var j=0; j<scoredElements.length; j++) {
                                optionPickone = new OptionElement( "", property );
                                optionPickone.appendChild( scoredElements[j] );
                                options.push( optionPickone );
                            }
                        }
                        else
                        {
                            for(var j=0; j<scoredElements.length; j++) {
                                option.appendChild( scoredElements[j] );
                            }
                            options.push( option );
                        }
                    }
                }
            }
        }
        else
        {
            for (var i = 0; i < pickoneList.length; i++)
            {
                var name = pickoneList[i];

                option = null;
                
                if( containsInArray( name, pickoneList ) )
                {
                    option = new OptionElement( name, property );
                    if( option ) {
                        options.push( option );
                    }
                }

                if ((option !== null) && (optionTable !== undefined))
                {
                    scoredElements = createScoredPropertyElement(optionTable[name], property, featureName);

                    if( scoredElements.length >= 2 &&
                        scoredElements[0].getName() == scoredElements[1].getName() )
                    {
                        options.pop();
                        for(var j=0; j<scoredElements.length; j++) {
                            optionPickone = new OptionElement( name, property );
                            optionPickone.appendChild( scoredElements[j] );
                            options.push( optionPickone );
                        }
                    }
                    else
                    {
                        for (var j = 0; j < scoredElements.length; j++)
                        {
                            option.appendChild(scoredElements[j]);
                        }
                    }
                }
            }
        }
    }
    catch(e) {
        options = [];
    }
        
    return options;
}

function createUserFormOptionElement( property, featureName, userformlist )
{
    var options = [];
    try
    {
        if( global.eventtype == EVENT_PRINTCAPABILITIES )
        {
            for(var userform in userformlist)
            {
                var option = new OptionElement( userform, property );
                var widthElement = new ScoredPropertyIntegerElement(
                                        USERFORM_SUPPORT_TABLE[featureName].width,
                                        userformlist[userform].width,
                                        property,
                                        featureName
                                        );

                var heightElement = new ScoredPropertyIntegerElement(
                                        USERFORM_SUPPORT_TABLE[featureName].height,
                                        userformlist[userform].height,
                                        property,
                                        featureName
                                        );

                option.appendChild( widthElement );
                option.appendChild( heightElement );

                options.push(option);
            }
        }
        else if( global.eventtype == EVENT_DEVMODETOPRINTTICKET )
        {
            var paper = property.getCurrentValue( featureName ); 
            if( paper.match("ns0000:UserForm") )
            {
                var option = new OptionElement( paper, property );
                var widthElement = new ScoredPropertyIntegerElement(
                                        USERFORM_SUPPORT_TABLE[featureName].width,
                                        property.getCurrentValue(USERFORM_SUPPORT_TABLE[featureName].width),
                                        property,
                                        featureName
                                        );

                var heightElement = new ScoredPropertyIntegerElement(
                                        USERFORM_SUPPORT_TABLE[featureName].height,
                                        property.getCurrentValue(USERFORM_SUPPORT_TABLE[featureName].height),
                                        property,
                                        featureName
                                        );

                option.appendChild( widthElement );
                option.appendChild( heightElement );

                options.push(option);
            }
        }
    }
    catch(e) {
        options = [];
    }
    return options;
}

function isBlankObject( obj )
{
    if(obj === undefined)
    {
        return false;
    }
    else if(obj.blank === undefined)
    {
        return false;
    }
    else
    {
        return true;
    }
}

function isArrayObject( obj )
{
    if( Object.prototype.toString.call(obj) === '[object Array]' )
    {
        return true;
    }
    else
    {
        return false;
    }
}

function isHashObject( obj )
{
    if( obj == null ) {
        return false;
    }

    if( Object.prototype.toString.call(obj) === '[object Object]' )
    {
        return true;
    }
    else
    {
        return false;
    }
}

function createScoredPropertyElement( scoredTable, property, featureName )
{
    var scoredElements = [];
    try
    {
        var isPickOneProperty = false;
        if (scoredTable.length > 2 && scoredTable[0][KEY_NAME] == scoredTable[1][KEY_NAME])
        {
            isPickOneProperty = true;
        }
        for(var i=0; i<scoredTable.length; i++ )
        {
            var scored = null;
            if( "ref"  === scoredTable[i][KEY_TYPE] )
            {
                if( property.isSupported( scoredTable[i][KEY_VALUE] ) )
                {
                    scored = new ScoredPropertyParameterRefElement(
                            scoredTable[i][KEY_NAME],
                            scoredTable[i][KEY_VALUE],
                            property,
                            featureName
                            );
                }
            }
            else if( "integer" === scoredTable[i][KEY_TYPE] )
            {
                if (isPickOneProperty )
                {
                    if( property.isSupported(scoredTable[i][KEY_NAME]) )
                    {
                        scored = new ScoredPropertyIntegerElement(
                                scoredTable[i][KEY_NAME],
                                scoredTable[i][KEY_VALUE],
                                property,
                                featureName
                                );
                    }
                }
                else
                {
                    scored = new ScoredPropertyIntegerElement(
                            scoredTable[i][KEY_NAME],
                            scoredTable[i][KEY_VALUE],
                            property,
                            featureName
                            );
                }
            }
            else if( "string" == scoredTable[i][KEY_TYPE] )
            {
                scored = new ScoredPropertyStringElement(
                        scoredTable[i][KEY_NAME],
                        scoredTable[i][KEY_VALUE],
                        property,
                        featureName
                        );
            }

            if( scored ) {
                scoredElements.push( scored );
            }
        }
    }
    catch(e) {
        scoredElements = [];
    }

    return scoredElements;
}

function createParameterDefElement( pcTable, property )
{
    var parameterDef = null;
    try
    {
        if( TAG_PARAMETERDEF == pcTable[KEY_TAG] && property.isSupported( pcTable[KEY_ATTRIBUTE] )  )
        {
            var type;
            var name = pcTable[KEY_ATTRIBUTE];
            var value, unittype, mandatory, min, max, defaultvalue, display, multiple;

            type      = pcTable[KEY_CHILDREN][DATATYPE_VALUE_NAME];
            unittype  = pcTable[KEY_CHILDREN][UNITTYPE_VALUE_NAME];
            mandatory = pcTable[KEY_CHILDREN][MANDATORY_VALUE_NAME];

            if( MIN_VALUE_NAME in pcTable[KEY_CHILDREN] ) {
                min       = pcTable[KEY_CHILDREN][MIN_VALUE_NAME];
            }
            else if( MIN_LENGTH_NAME in pcTable[KEY_CHILDREN] ) {
                min       = pcTable[KEY_CHILDREN][MIN_LENGTH_NAME];
            }

            max = property.getMaxValue(name);
            if( !max )
            {
                if( MAX_VALUE_NAME in pcTable[KEY_CHILDREN] ) {
                    max       = pcTable[KEY_CHILDREN][MAX_VALUE_NAME];
                }
                else if( MAX_LENGTH_NAME in pcTable[KEY_CHILDREN] ) {
                    max       = pcTable[KEY_CHILDREN][MAX_LENGTH_NAME];
                }
            }

            if( MULTIPLE_VALUE_NAME in pcTable[KEY_CHILDREN] ) {
                multiple  = pcTable[KEY_CHILDREN][MULTIPLE_VALUE_NAME];
            }

            if( DISPLAYNAME_VALUE_NAME in pcTable[KEY_CHILDREN] ) {
                display   = pcTable[KEY_CHILDREN][DISPLAYNAME_VALUE_NAME];
            }

            defaultvalue = property.getDefault( name );
            if( defaultvalue == null && DEFAULTVAL_VALUE_NAME in pcTable[KEY_CHILDREN] )
            {
                defaultvalue = pcTable[KEY_CHILDREN][DEFAULTVAL_VALUE_NAME];
            }
            value = property.getCurrentValue( name );

            if( type == "xsd:integer" ) {
                parameterDef = new ParameterDefIntegerElement(
                                    property,
                                    name,
                                    value,
                                    unittype,
                                    mandatory,
                                    min,
                                    max,
                                    defaultvalue,
                                    multiple,
                                    display
                                    );
            }
            else if( type == "xsd:string" ) {
                parameterDef = new ParameterDefStringElement(
                                    property,
                                    name,
                                    value,
                                    unittype,
                                    mandatory,
                                    min,
                                    max,
                                    defaultvalue,
                                    display
                                    );
            }
        }
    }
    catch(e)
    {
        parameterDef = null;
    }
    return parameterDef;
}


var FeatureElement = function( property, name, parentFeatureName, display )
{
    this.property = property;
    this.name = name;
    this.parentFeatureName = parentFeatureName;
    this.display = display;
    this.children = [];
    this.childFeature = [];
};

FeatureElement.prototype.getName = function() {
    return this.name;
};

FeatureElement.prototype.appendChild = function( element )
{
    this.children.push( element );
};

FeatureElement.prototype.appendFeature = function( element )
{
    this.childFeature.push( element );
};

FeatureElement.prototype.createPrintTicket = function(rootNode, parentNode){
    //debugger;
    var ret = false;
    try
    {
        var node = null;
        if( (!(containsInArray( this.name, AUTO_CREATE_FEATURE_LIST )) && !("psk:PageResolution" == this.name) ) ||
            (("psk:PageResolution" == this.name) && this.property.devmode[this.name] === undefined ) ||
            global.isDefaultRequest
            )
        {
            removeChildNode( rootNode, "psf:PrintTicket/psf:Feature[@name='" + this.name + "']" );

            node = CreateFeature(
                            rootNode,
                            this.name.substring( this.name.indexOf(":") + 1),
                            null,  // displayName
                            this.name.substring( 0, this.name.indexOf(":") )
                            );

            if( node )
            {
                var value = this.property.getCurrentValue( this.name, this.parentFeatureName );
                for(var i=0; i < this.children.length; i++ )
                {
                    if( value == this.children[i].getName() ||
                        ( value === null && this.children[i].getName() === "" )
                        )
                    {
                        ret = this.children[i].createPrintTicket(rootNode, node );
                        if( ret ) {
                            break;
                        }
                    }
                }
            }
        }
        else
        {
            node = rootNode.XmlNode.selectSingleNode("psf:PrintTicket/psf:Feature[@name='" + this.name + "']");
        }

        for(var i=0; i < this.childFeature.length; i++ )
        {
            var retchild = this.childFeature[i].createPrintTicket(rootNode, node );
            ret = (!ret && retchild) ? retchild : ret;
        }

        if( ret )
        {
            parentNode.appendChild( node );
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

FeatureElement.prototype.createPrintCapabilities = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var node = null;
        if( !(containsInArray( this.name, AUTO_CREATE_FEATURE_LIST ) ) )
        {
            removeChildNode( rootNode, "psf:PrintCapabilities/psf:Feature[@name='" + this.name + "']" );

            node = CreateFeatureSelection(
                            rootNode,
                            this.name.substring( this.name.indexOf(":") + 1),
                            null,
                            this.name.substring( 0, this.name.indexOf(":") )
                            );

            if( node )
            {
                for(var i=0; i < this.children.length; i++ )
                {
                    ret = this.children[i].createPrintCapabilities(rootNode, node );
                }
            }
        }
        else
        {
            node = rootNode.XmlNode.selectSingleNode("psf:PrintCapabilities/psf:Feature[@name='" + this.name + "']");
        }

        if( node )
        {
            for(var i=0; i < this.childFeature.length; i++ )
            {
                ret = this.childFeature[i].createPrintCapabilities(rootNode, node );
            }
        }

        if( ret )
        {
            parentNode.appendChild( node );
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

var OptionElement = function( name, property, display )
{
    this.name = name;
    this.property = property;
    this.display = display;
    this.children = [];
};

OptionElement.prototype.getName = function()
{
    return this.name;
};

OptionElement.prototype.appendChild = function( element )
{
    this.children.push( element );
};

OptionElement.prototype.createPrintTicket = function(rootNode, parentNode ){
    var ret = false;
    try
    {
        var node = CreateOption(
                        rootNode,
                        this.name.substring( this.name.indexOf(":") + 1),
                        null, 
                        this.name.substring( 0, this.name.indexOf(":") )
                        );

        if( node )
        {
            ret = true;
            if( this.children.length > 0 &&
                this.property.getPickoneList( this.children[0].getName() ) !== null )
            {
                for(var i=0; i < this.children.length; i++ )
                {
                    ret = this.children[i].createPrintTicket(rootNode, node );
                    if( ret )
                    {
                        break;
                    }
                }

            }
            else
            {
                for(var i=0; i < this.children.length; i++ )
                {
                    ret = this.children[i].createPrintTicket(rootNode, node );
                }
            }
        }

        if( ret )
        {
            parentNode.appendChild( node );
        }

    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

OptionElement.prototype.createPrintCapabilities = function(rootNode, parentNode ){
    var ret = false;
    try
    {
        var node = CreateOption(
                        rootNode,
                        this.name.substring( this.name.indexOf(":") + 1),
                        null, 
                        this.name.substring( 0, this.name.indexOf(":") )
                        );

        if( node )
        {
            ret = true;
            for(var i=0; i < this.children.length; i++ )
            {
                ret = this.children[i].createPrintCapabilities(rootNode, node );
            }
        }

        if( ret )
        {
            parentNode.appendChild( node );
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

var ScoredPropertyIntegerElement = function( name, value, property, featureName)
{
    this.name = name;
    this.value = value;
    this.property = property;
    this.featureName = featureName;
};

var ScoredPropertyStringElement = function( name, value, property, featureName)
{
    this.name = name;
    this.value = value;
    this.property = property;
    this.featureName = featureName;
};

var ScoredPropertyParameterRefElement = function( name, refName, property, featureName )
{
    this.name = name;
    this.ref = refName;
    this.property = property;
    this.featureName = featureName;
};

ScoredPropertyParameterRefElement.prototype.createPrintTicket = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var node = CreateScoredPropertyRefElement(
                        rootNode,
                        this.name.substring( this.name.indexOf(":") + 1),
                        this.ref.substring( this.ref.indexOf(":") + 1),
                        this.name.substring( 0, this.name.indexOf(":") )
                        );

        if( node )
        {
            this.property.setReferedParameterDef( this.ref );
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};


ScoredPropertyParameterRefElement.prototype.createPrintCapabilities = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var node = CreateScoredPropertyRefElement(
                        rootNode,
                        this.name.substring( this.name.indexOf(":") + 1),
                        this.ref.substring( this.ref.indexOf(":") + 1),
                        this.name.substring( 0, this.name.indexOf(":") )
                        );

        if( node )
        {
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};


ScoredPropertyParameterRefElement.prototype.getRefName = function() {
    return this.ref;
};

ScoredPropertyParameterRefElement.prototype.getName = function() {
    return this.name;
};

ScoredPropertyIntegerElement.prototype.createPrintTicket = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var value = this.property.getCurrentValue( this.name, this.featureName );
        var node = null;
        if( (! this.property.isSupported(this.name) && value === null ) || value == this.value )
        {
            node = CreateScoredPropertyIntegerElement(
                            rootNode,
                            this.name.substring( this.name.indexOf(":") + 1),
                            this.value,
                            this.name.substring( 0, this.name.indexOf(":") )
                            );
        }

        if( node )
        {
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};


ScoredPropertyIntegerElement.prototype.createPrintCapabilities = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var node = CreateScoredPropertyIntegerElement(
                        rootNode,
                        this.name.substring( this.name.indexOf(":") + 1),
                        this.value,
                        this.name.substring( 0, this.name.indexOf(":") )
                        );

        if( node )
        {
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};


ScoredPropertyIntegerElement.prototype.getName = function() {
    return this.name;
};


ScoredPropertyStringElement.prototype.createPrintTicket = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var value = this.property.getCurrentValue( this.name, this.featureName );
        var node = null;
        if( value === null || value == this.value )
        {
            node = CreateScoredPropertyStringElement(
                            rootNode,
                            this.name.substring( this.name.indexOf(":") + 1),
                            this.value,
                            this.name.substring( 0, this.name.indexOf(":") )
                            );
        }

        if( node )
        {
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};


ScoredPropertyStringElement.prototype.createPrintCapabilities = function(rootNode, parentNode){
    var ret = false;
    try
    {
        var node = CreateScoredPropertyStringElement(
                        rootNode,
                        this.name.substring( this.name.indexOf(":") + 1),
                        this.value,
                        this.name.substring( 0, this.name.indexOf(":") )
                        );

        if( node )
        {
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

ScoredPropertyStringElement.prototype.getName = function() {
    return this.name;
};



var ParameterDefIntegerElement = function( property, name, value, unittype, mandatory, min, max, defaultvalue, multiple, display)
{
    this.property= property;
    this.name = name;
    this.value = value;
    this.unittype = unittype;
    this.mandatory = mandatory;
    this.min = min;
    this.max = max;
    this.defaultvalue = (defaultvalue === null) ? min : defaultvalue;
    this.multiple = multiple;
    this.display = display;
};

ParameterDefIntegerElement.prototype.createPrintTicket = function(rootNode, parentNode){
    var ret = false;
    try
    {
        if( !containsInArray(this.name, AUTO_CREATE_PARAMETERDEF_LIST) || global.isDefaultRequest )
        {
            removeChildNode( rootNode, "psf:PrintTicket/psf:ParameterInit[@name='" + this.name + "']" );

            var node = CreateIntParamRefIni(
                    rootNode,
                    this.name.substring( this.name.indexOf(":") + 1),
                    (this.name.substring( 0, this.name.indexOf(":") ) == "psk" ),
                    "xsd:integer",
                    (this.value == null && this.property.isReferedParameterDef(this.name)) ? this.defaultvalue : this.value
                    );

            if( node )
            {
                parentNode.appendChild( node );
                ret = true;
            }
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

ParameterDefIntegerElement.prototype.createPrintCapabilities = function(rootNode, parentNode){
    var ret = false;
    try
    {
        if (!(containsInArray(this.name, AUTO_CREATE_PARAMETERDEF_LIST))) {
            removeChildNode(rootNode, "psf:PrintCapabilities/psf:ParameterDef[@name='" + this.name + "']");

            var node = CreateIntParamDef(
                  rootNode,
                  this.name.substring(this.name.indexOf(":") + 1),
                  (this.name.substring(0, this.name.indexOf(":")) == "psk"),
                  this.defaultvalue,
                  this.min,
                  this.max,
                  this.multiple,
                  this.unittype
                  );

            if (node) {
                parentNode.appendChild(node);
                ret = true;
            }
        }
        else {
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

var ParameterDefStringElement = function( property, name, value, unittype, mandatory, min, max, defaultvalue, display)
{
    this.property= property;
    this.name = name;
    this.value = value;
    this.unittype = unittype;
    this.mandatory = mandatory;
    this.min = min;
    this.max = max;
    this.defaultvalue = (defaultvalue == null) ? "" : defaultvalue;
    this.display = display;
};

ParameterDefStringElement.prototype.createPrintTicket = function(rootNode, parentNode){
    //debugger;
    var ret = false;
    try
    {
        removeChildNode( rootNode, "psf:PrintTicket/psf:ParameterInit[@name='" + this.name + "']" );

        {
            var node = CreateIntParamRefIni(
                    rootNode,
                    this.name.substring( this.name.indexOf(":") + 1),
                    (this.name.substring( 0, this.name.indexOf(":") ) == "psk" ),
                    "xsd:string",
                    (this.value == null && this.property.isReferedParameterDef(this.name)) ? this.defaultvalue : this.value
                    );

            if( node )
            {
                parentNode.appendChild( node );
                ret = true;
            }
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};

ParameterDefStringElement.prototype.createPrintCapabilities = function(rootNode, parentNode){
    var ret = false;
    try
    {
        removeChildNode( rootNode, "psf:PrintCapabilities/psf:ParameterDef[@name='" + this.name + "']" );

        var node = CreateStringParamDef(
                rootNode,
                this.name.substring( this.name.indexOf(":") + 1),
                (this.name.substring( 0, this.name.indexOf(":") ) == "psk" ),
                this.defaultvalue,
                this.min,
                this.max,
                this.unittype
                );

        if( node ) {
            parentNode.appendChild( node );
            ret = true;
        }
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
};


var ModelSpec = function( modelSpecJson ) {
    this.modelSpec = modelSpecJson;
    this.initialize();

    this.getDefault = function( name ) {
        if( name in this.modelSpec && KEY_DEFAULT in this.modelSpec[name] )
        {
            return this.modelSpec[name][KEY_DEFAULT];
        }
        else
        {
            return null;
        }
    };

    this.getPickoneList = function( name ) {
        if( name in this.modelSpec )
        {
            return this.modelSpec[name][KEY_PICKONE];
        }
        else
        {
            return null;
        }
    };

    this.getMinValue = function( name ) {
        return this.modelSpec[name][KEY_MIN];
    };

    this.getMaxValue = function( name ) {
        return this.modelSpec[name][KEY_MAX];
    };

    this.isSupported = function( name ) {
        if( name in this.modelSpec ||
            containsInArray(name, AUTO_CREATE_FEATURE_LIST) ||
            containsInArray(name, AUTO_CREATE_PARAMETERDEF_LIST) )
        {
            return true;
        }
        else
        {
            return false;
        }
    };

    this.getFeatureNameList = function() {
        var list = [];
        for( var feature in this.modelSpec ) {
            list.push( feature );
        }
        for( var feature in USERFORM_SUPPORT_TABLE )
        {
            for( var wh in USERFORM_SUPPORT_TABLE[feature] )
            {
                list.push( USERFORM_SUPPORT_TABLE[feature][wh]  );
            }
        }
        return list;
    };

};

ModelSpec.prototype.initialize = function() {
    for(var feature in RELATED_FEATURE_LIST )
    {
        if( feature in this.modelSpec )
        {
            if( "name" in RELATED_FEATURE_LIST[feature] &&
                "default" in RELATED_FEATURE_LIST[feature] )
            {
                this.modelSpec[RELATED_FEATURE_LIST[feature]["name"]] = {"default":RELATED_FEATURE_LIST[feature]["default"] };
                if( "pickone" in RELATED_FEATURE_LIST[feature] )
                {
                    this.modelSpec[RELATED_FEATURE_LIST[feature]["name"]]["pickone"] = RELATED_FEATURE_LIST[feature]["pickone"];
                }
            }
        }
    }
};



function removeChildNode( ptpc, xpath )
{
    var node = ptpc.XmlNode.selectSingleNode(xpath);
    if( node ) {
        ptpc.XmlNode.documentElement.removeChild( node );
    }
}

var Property = function( modelSpec, devmode, rcf, ptpsdMap ) {
    this.modelSpec = modelSpec;
    this.devmode   = devmode;
    this.rcf       = rcf;
    this.ptpsdMap  = ptpsdMap;
    this.refered   = [];

    // featureName is Optional augment
    this.getCurrentValue = function( name, featureName )
    {
        var result = null;
        if( this.rcf )
        {
            var rcfValue = [];
            var psdkey = PtToPsdFeatureKey(name, this.ptpsdMap);
            if( (psdkey !== null && psdkey.length > 0) )
            {
                var fixvalue = this.getFixValueToArray( psdkey );
                if(fixvalue.length > 0)
                {
                    rcfValue = PsdToPtValueKey( name, fixvalue, this.ptpsdMap );
                    if( rcfValue.length > 0 )
                    {
                        result = rcfValue[0];
                    }
                }

                if( !rcfValue.length > 0 &&
                    global.eventtype === EVENT_DEVMODETOPRINTTICKET &&
                    global.isDefaultRequest )
                {
                    var rcfDefault = this.getDefaultValueToArray( psdkey );
                    if(rcfDefault.length > 0)
                    {
                        var getDefault = PsdToPtValueKey( name, rcfDefault, this.ptpsdMap );
                        if( getDefault.length > 0 )
                        {
                            result = getDefault[0];
                        }
                    }
                }
            }
        }

        if( result === null )
        {
            if( this.devmode && name in this.devmode && !global.isDefaultRequest )
            {
                if( featureName != null &&
                    containsInArray(name, DUPLICATED_FEATURENAME_LIST) &&
                    isHashObject(this.devmode[name])
                    )
                {
                    if( featureName in this.devmode[name] )
                    {
                        result = this.devmode[name][featureName];
                    }
                    else
                    {
                        result = this.modelSpec.getDefault(name);
                    }
                }
                else
                {
                    result = this.devmode[name];
                }
            }
            else
            {
                result = this.modelSpec.getDefault(name);
            }
        }

        if( result )
        {
            var pickoneList = this.getPickoneList( name );
            if( pickoneList &&
                !containsInArray(result, pickoneList) &&
                !containsInArray(name, CHANGEABLE_PICKONE_FEATURE_LIST ) &&
                !(name in USERFORM_SUPPORT_TABLE) )
            {
                result = this.modelSpec.getDefault( name );
                if (!containsInArray(result, pickoneList) )
                {
                    result = pickoneList[0];
                }
            }
        }

        return result;
    };

    // featureName is optional argument
    this.setCurrentValue = function( name, value, featureName ) {
        var result = false;
        if( value != null )
        {
            if( this.rcf )
            {
                var psdkey = PtToPsdFeatureKey(name, this.ptpsdMap);
                if( (psdkey !== null && psdkey.length > 0) )
                {
                    var fixvalue = [];
                    fixvalue = this.getFixValueToArray( psdkey );
                    if(fixvalue.length > 0)
                    {
                        var currentValue = PsdToPtValueKey( name, fixvalue, this.ptpsdMap );
                        if( currentValue.length > 0 )
                        {
                            this.devmode[name] = currentValue[0];
                            result = true;
                        }
                    }
                }
            }

            if( !result ) {
                if( featureName != null && containsInArray(name, DUPLICATED_FEATURENAME_LIST) )
                {
                    if( this.devmode[name] == null || !isHashObject(this.devmode[name]) )
                    {
                        this.devmode[name] = {};
                    }
                    this.devmode[name][featureName] = value;
                }
                else
                {
                    this.devmode[name] = value;
                }
            }
        }
    };

    this.getDefault = function( name )
    {
        var rcfDefault = this.getCustomizedDefault( name );

        return (rcfDefault) ? rcfDefault : this.modelSpec.getDefault(name);
    };

    this.getCustomizedDefault = function( name )
    {
        if( this.rcf )
        {
            var psdkey = PtToPsdFeatureKey(name, this.ptpsdMap);
            if( (psdkey !== null && psdkey.length > 0) )
            {
                var customDefault = this.getDefaultValueToArray( psdkey );
                if(customDefault.length > 0)
                {
                    var getDefault = PsdToPtValueKey( name, customDefault, this.ptpsdMap );
                    if( getDefault.length > 0 )
                    {
                        return getDefault[0];
                    }
                }
            }
        }
        return null;
    };

    this.getPickoneList = function( name )
    {
        if( this.rcf )
        {
            var psdKeyName = PtToPsdFeatureKey(name, this.ptpsdMap);
            if( (psdKeyName !== null && psdKeyName.length === 1) )
            {
                var pickoneList = [];
                for ( var i=0; i<psdKeyName.length; i++ )
                {
                    var getValue = rcf.getPickoneList( psdKeyName[i] );
                    if( getValue !== null )
                    {
                        pickoneList.push(getValue);
                    }
                }
                if( pickoneList.length > 0 )
                {
                    var rcfpickone = ChangePickoneList( pickoneList );
                    var pickone = [];
                    for(var j=0; j<rcfpickone.length; j++ )
                    {
                        var getValue = PsdToPtValueKey( name, rcfpickone[j], this.ptpsdMap );
                        for (var k=0; k<getValue.length; k++ )
                        {
                            if( !(containsInArray( getValue[k].toString(), pickone ) ) )
                            {
                                pickone.push( getValue[k].toString() );
                            }
                        }
                    }
                    if( pickone.length > 0 )
                    {
                        return pickone;
                    }
                }

                var fixvalue = this.getFixValueToArray( psdKeyName );
                if(fixvalue.length > 0)
                {
                    return PsdToPtValueKey( name, fixvalue, this.ptpsdMap );
                }
            }
        }

        return this.modelSpec.getPickoneList(name);
    };

    this.getDevmodeValue = function( name )
    {
        if( (this.devmode != null) && (name in this.devmode) && !isHashObject(this.devmode[name]) )
        {
            return this.devmode[name];
        }
        else
        {
            return null;
        }
    }

    this.getMinValue = function( name ) {
        return this.modelSpec.getMinValue(name);
    };

    this.getFixValue = function( name ) {
        var result = null;
        if( this.rcf )
        {
            var psdkey = PtToPsdFeatureKey(name, this.ptpsdMap);
            if( (psdkey !== null && psdkey.length > 0) )
            {
                var fixvalue = this.getFixValueToArray( psdkey );
                if(fixvalue.length > 0)
                {
                    var ptValue = PsdToPtValueKey( name, fixvalue, this.ptpsdMap );
                    if( ptValue.length > 0 )
                    {
                        result = ptValue[0];
                    }
                }
            }
        }
        return result;
    };

    this.getFixValueToArray = function( psdkey ) {
        var result = [];
        var bool = false;
        for( var i=0; i<psdkey.length; i++ )
        {
            var getValue = this.rcf.getFixValue( psdkey[i] );
            result.push( getValue );
        }
        for( var j=0; j<result.length; j++ )
        {
            if(result[j] != null)
            {
                bool = true;
                break;
            }
        }
        if(!bool)
        {
            result = [];
        }

        return result;
    };

    this.getDefaultValueToArray = function( psdkey ) {
        var result = [];
        var bool = false;
        for( var i=0; i<psdkey.length; i++ )
        {
            var getValue = this.rcf.getDefault( psdkey[i] );
            result.push( getValue );
        }
        for( var j=0; j<result.length; j++ )
        {
            if(result[j] != null)
            {
                bool = true;
                break;
            }
        }
        if(!bool)
        {
            result = [];
        }

        return result;
    };

    this.getMaxValue = function( name ) {
        var result = null;
        if( this.rcf )
        {
            var psdName = PtToPsdFeatureKey(name, this.ptpsdMap);
            if( (psdName !== null && psdName.length > 0) )
            {
                for( var i=0; i<psdName.length; i++ )
                {
                    result = this.rcf.getMaxValue( psdName[i] );
                }
            }
        }
        
        if( !result )
        {
            result = this.modelSpec.getMaxValue(name);
        }
        return result;
    };

    this.isSupported = function( name ) {
        return this.modelSpec.isSupported(name);
    };

    this.isSupportedElement = function( feature, value ) {
        var result = false;
        if( containsInArray(value, this.getPickoneList(feature)) )
        {
            result = true;
        }
        return result;
    };

    this.getFeatureNameList = function() {
        return this.modelSpec.getFeatureNameList();
    };

    this.setReferedParameterDef = function( name ) {
        return this.refered.push(name);
    };

    this.isReferedParameterDef = function( name ) {
        return containsInArray( name, this.refered );
    };

    this.setLandscapeDirection = function( value ) {
        if( this.isSupportedElement(FEATURE_PAGE_ORIENTATION, OPTION_REVERSELANDSCAPE) )
        {
            if( value == OPTION_LANDSCAPE || value == OPTION_REVERSELANDSCAPE )
            {
                this.devmode[FEATURE_PAGE_ORIENTATION] = value;
            }
        }
    };

    this.validate = function() {
        try
        {
            this.inheritValue(
                 PRINT_QUALITY_PRESET,
                 PRINT_QUALITY_ADJUSTMENT,
                 REPLACE_MAP_PRINT_QUALITY_PRESET
                 );

            this.inheritValue(
                 COLOR_PROFILE,
                 COLOR_PROFILE_SETTING,
                 REPLACE_MAP_COLOR_PROFILE
                 );

            this.inheritValue(
                 DITHER,
                 DITHER_SETTING,
                 REPLACE_MAP_DITHER
                 );
        }
        catch(e)
        {
        }
    };

    this.inheritValue = function( oldFeature, newFeature, replaceMap )
    {
        try
        {
            var supportFeatureType = "notsupported"; 
            if( this.isSupported(newFeature) )
            {
                supportFeatureType = (this.isSupported(oldFeature)) ? "oldandnew" : "new";
            }
            else
            {
                supportFeatureType = (this.isSupported(oldFeature)) ? "old" : "notsupported";
            }

            var devmodeType = "notsupported"; 
            if( newFeature in this.devmode )
            {
                devmodeType = (oldFeature in this.devmode) ? "oldandnew" : "new";
            }
            else
            {
                devmodeType = (oldFeature in this.devmode) ? "old" : "notsupported";
            }


            if( supportFeatureType === "oldandnew" && devmodeType === "old" )
            {
                if( this.devmode[oldFeature] in replaceMap )
                {
                    this.devmode[newFeature] = replaceMap[this.devmode[oldFeature]];
                }
            }
            else if( supportFeatureType === "old" && devmodeType === "oldandnew" )
            {
                var convertedValue = GetKeyFromValue(this.devmode[newFeature], replaceMap);
                if( convertedValue != null )
                {
                    this.devmode[oldFeature] = convertedValue;
                    delete this.devmode[newFeature];
                }
            }
            else if( supportFeatureType === "new" && devmodeType === "old" )
            {
                if( this.devmode[oldFeature] in replaceMap )
                {
                    this.devmode[newFeature] = replaceMap[this.devmode[oldFeature]];
                    delete this.devmode[oldFeature];
                }
            }
            else if( supportFeatureType === "old" && devmodeType === "new" )
            {
                var convertedValue = GetKeyFromValue(this.devmode[newFeature], replaceMap);
                if( convertedValue !== null )
                {
                    this.devmode[oldFeature] = convertedValue;
                    delete this.devmode[newFeature];
                }
            }
            else
            {
                // no change
            }
        }
        catch(e)
        {
        }
    };
};

function GetKeyFromValue( value, hash )
{
    for (var key in  hash)
    {
        if (value === hash[key])
        {
            return key;
        }
    }
    return null;
}

//
// RCFの処理に関する
//
var RCF = function( rcfDom, ptpsdMap ) {
    this.dom       = rcfDom;
    this.rcfInfo   = {};
    this.initialize();
};

RCF.prototype.initialize = function()
{
    try
    {
        var version = RCFParser_GetVersion( this.dom );
        
        if( version == "1.0" )
        {
            var customer       = RCFParser_GetCustomer( this.dom );
            var devicesettings = RCFParser_GetDeviceSettings( this.dom);
            var featurelock    = RCFParser_GetFeatureLock( this.dom );
            var modifyitem     = RCFParser_GetModifyItem( this.dom );

            var item = {};
            if( devicesettings[RCF_ITEM_KEY_ITEMLIST] )
            {
                for(var i=0; i<devicesettings[RCF_ITEM_KEY_ITEMLIST].length; i++ )
                {
                    item = {};
                    item[RCF_ITEM_KEY_DEFAULT] = devicesettings[RCF_ITEM_KEY_ITEMLIST][i][RCF_ITEM_KEY_VALUE];
                    this.rcfInfo[ devicesettings[RCF_ITEM_KEY_ITEMLIST][i][RCF_ITEM_KEY_NAME] ] = item;
                }
            }

            for(var i=0; i<featurelock.length; i++ )
            {
                item = {};
                if( featurelock[i][RCF_ITEM_KEY_NAME] in this.rcfInfo )
                {
                    item = this.rcfInfo[ featurelock[i][RCF_ITEM_KEY_NAME] ];
                }
                item[RCF_ITEM_KEY_FIXVALUE] = featurelock[i][RCF_ITEM_KEY_FIXVALUE];
                item[RCF_ITEM_KEY_PICKONE]  = featurelock[i][RCF_ITEM_KEY_PICKONE];

                this.rcfInfo[ featurelock[i][RCF_ITEM_KEY_NAME] ] = item;
            }
            
            for(var i=0; i<modifyitem.length; i++)
            {
                item = {};
                if( modifyitem[i][RCF_ITEM_KEY_NAME] in this.rcfInfo )
                {
                    item = this.rcfInfo[ modifyitem[i][RCF_ITEM_KEY_NAME] ];
                }
                item[RCF_ITEM_KEY_MAX] = modifyitem[i][RCF_ITEM_KEY_MAX];

                this.rcfInfo[ modifyitem[i][RCF_ITEM_KEY_NAME] ] = item;
            }


        }
    }
    catch(e)
    {
        this.rcfInfo = {};
    }
};

RCF.prototype.getDefault = function( name )
{
    var result = null;
    if( name in this.rcfInfo )
    {
        if( RCF_ITEM_KEY_DEFAULT in this.rcfInfo[name] )
        {
            result = this.rcfInfo[name][RCF_ITEM_KEY_DEFAULT];
        }
        else if( RCF_ITEM_KEY_FIXVALUE in this.rcfInfo[name] )
        {
            result = rcfInfo[name][RCF_ITEM_KEY_FIXVALUE];
        }
        else if( RCF_ITEM_KEY_PICKONE in this.rcfInfo[name] &&
                 this.rcfInfo[name][RCF_ITEM_KEY_PICKONE].length == 1 )
        {
            result = this.rcfInfo[name][RCF_ITEM_KEY_FIXVALUE][0];
        }
        else
        {
            result = null;
        }
    }
    return result;
};

RCF.prototype.getFixValue = function( name )
{
    var result = null;
    if( name in this.rcfInfo )
    {
        if( RCF_ITEM_KEY_FIXVALUE in this.rcfInfo[name] )
        {
            result = this.rcfInfo[name][RCF_ITEM_KEY_FIXVALUE];
        }
        else if(    RCF_ITEM_KEY_PICKONE in this.rcfInfo[name] &&
                    this.rcfInfo[name][RCF_ITEM_KEY_PICKONE].length == 1 )
        {
            result = this.rcfInfo[name][RCF_ITEM_KEY_PICKONE][0];
        }
        else
        {
            result = null;
        }
    }
    return result;
};

RCF.prototype.getPickoneList = function( name )
{
    var result = null;
    if(    name in this.rcfInfo && RCF_ITEM_KEY_PICKONE in this.rcfInfo[name] )
    {
        result = this.rcfInfo[name][RCF_ITEM_KEY_PICKONE];
    }
    else
    {
        result = null;
    }
    return result;
};

RCF.prototype.getMaxValue = function( name )
{
    var result = null;
    if( name in this.rcfInfo )
    {
        if( RCF_ITEM_KEY_MAX in this.rcfInfo[name] )
        {
            result = this.rcfInfo[name][RCF_ITEM_KEY_MAX];
        }
    }
    return result;
};

function ChangePickoneList( pickoneList )
{
    var result = [];
    var index = 0;

    if( pickoneList[0] !== undefined )
    {
        for (var i =0; i<pickoneList[0].length; i++ )
        {
            if( pickoneList[1] !== undefined )
            {
                for (var j=0; j<pickoneList[1].length; j++ )
                {
                    result[index] = [];
                    result[index][0] = pickoneList[0][i];
                    result[index][1] = pickoneList[1][j];
                    index++;
                }
            }
            else
            {
                result[index] = [];
                result[index][0] = pickoneList[0][i];
                index++;
            }
        }
    }
    return result;
}

function getDevmodeObject( devModeProperties )
{
    var settingsObject = null;
    try
    {
        var settings = devModeProperties.GetString("DocumentSettings");

        settingsObject = (new Function("return " + settings))();
    }
    catch(e)
    {
        settingsObject = null;
    }

    return settingsObject;
}


function getInnerSpecObject( scriptContext )
{
    var innerSpec = scriptContext.DriverProperties.GetString("pc_table");

    var innerSpecObject = ConvertToJsonObjectFromString(innerSpec);

    return innerSpecObject;
}


function getModelSpecObject( scriptContext, modelName )
{
    var modelSpec = "";
    if ( 0 === modelName.indexOf("Generic") )
    {
        modelSpec = GetModelSpecTable(scriptContext, modelName);
        modelSpec = setAdditionalFeatureString_Generic(modelSpec);
    }
    else
    {
        modelSpec = GetModelSpecTableFromStream(scriptContext, modelName);
        modelSpec = setAdditionalFeatureString(modelSpec);
    }

    var modelSpecJson = ConvertToJsonObjectFromString(modelSpec);

    var modelSpecObject = new ModelSpec( modelSpecJson );
    //setPageProfileDistributionDefault(modelSpecObject, scriptContext);

    return modelSpecObject;
}


function getXPathTableObject(scriptContext) {

    var str = scriptContext.DriverProperties.GetString("pt_xpath");
    var xpathTable = ConvertToJsonObjectFromString(str);

    return xpathTable;
}


function getPtPsdMapObject( scriptContext )
{
    var pt_table = scriptContext.DriverProperties.GetString("pt_table");
    var tableObj = ConvertToJsonObjectFromString(pt_table);

    return tableObj;
}


function createAllPrintTicket(devModeProperties, scriptContext, printTicket)
{
    try {
        var modelName = GetCurrentModelName(scriptContext);
        if( modelName !== null )
        {
            var property = createPropertyObject(scriptContext, devModeProperties, printTicket );

            var innerSpecObject = getInnerSpecObject( scriptContext );

            innerSpecObject = addPagePrintPaperSizeCurrentUserForm(property, innerSpecObject);

            var table = getPrintTicketTreeTable( innerSpecObject, property, scriptContext, null );
            for( var i=0; i<table.length; i++ )
            {
                table[i].createPrintTicket(
                            printTicket, 
                            printTicket.XmlNode.lastChild
                            );
            }

            updateResolutionDevmodeValue( printTicket, property);

            updateReverseLandscapePrintTicket( printTicket, property);
        }
    }
    catch (e) {
        var result = 0;
        result = e.number;
        Debug.write(e.message + "createAllPrintTicket() Error!");
    }
}

function addPagePrintPaperSizeCurrentUserForm(property, innerSpecObject)
{
    try {
        var printSizeName = property.getCurrentValue( "ns0000:PagePrintPaperSize" );
        if( (printSizeName !== null) &&
            (printSizeName.match("ns0000:UserForm")) )
        {
            var pickoneList = createUserFormPickOne(property);
            for (var i = 0; i < innerSpecObject.length; i++)
            {
                if (innerSpecObject[i].attribute == "ns0000:PagePrintPaperSize")
                {
                    innerSpecObject[i].pickone[printSizeName] = {};
                    innerSpecObject[i].pickone[printSizeName] = pickoneList;
                    break;
                }
            }
        }
    }
    catch (e) {
        Debug.write(e.message + "addPagePrintPaperSizeCurrentUserForm() Error!");
    }

    return innerSpecObject;
}

function createUserFormPickOne(property)
{
    var pickoneList = [];

    try
    {
        pickoneList[0] = {
            "name": "ns0000:MediaSizeWidth",
            "type": TYPE_INTEGER,
            "value": property.getCurrentValue(CHANGFEATURE_TABLE_LIST[0])
        };
        pickoneList[1] = {
            "name": "ns0000:MediaSizeHeight",
            "type": TYPE_INTEGER,
            "value": property.getCurrentValue(CHANGFEATURE_TABLE_LIST[1])
        };
    }
    catch (e)
    {
        Debug.write(e.message + "createUserFormPickOne() Error!");
    }

    return pickoneList;
}

function updateResolutionDevmodeValue(printTicket, property) {
    var xpath = null;

    try
    {
        xpath = "psf:Feature[@name='psk:PageResolution']/psf:Option";
        var resolution = printTicket.XmlNode.documentElement.selectSingleNode(xpath);

        if ( (resolution) &&
             (resolution.childNodes.length != 0) &&
             (global.isDefaultRequest  !== true)
             )
        {
            if ((RESOLUTION_MAP[property.devmode['psk:PageResolution']]['ResolutionX'] == resolution.childNodes[0].text) && (RESOLUTION_MAP[property.devmode['psk:PageResolution']]['ResolutionY'] == resolution.childNodes[1].text))
            {
                resolution.attributes[0].text = property.devmode['psk:PageResolution'];
            }
        }
    }
    catch (e)
    {
        Debug.write(e.message + "updateResolutionDevmodeValue() Error!");
    }
}

function updateReverseLandscapePrintTicket(printTicket, property) {
    var xpath = null;

    try
    {
        if( property.isSupportedElement(FEATURE_PAGE_ORIENTATION, OPTION_REVERSELANDSCAPE) )
        {
            var xpath = "psf:Feature[@name= '" + FEATURE_PAGE_ORIENTATION + "']/psf:Option[@name='" + OPTION_LANDSCAPE + "']";
            var node = printTicket.XmlNode.documentElement.selectSingleNode( xpath );
            if( node )
            {
                if( property.getCurrentValue(FEATURE_PAGE_ORIENTATION) == OPTION_REVERSELANDSCAPE )
                {
                    node.attributes[0].text = OPTION_REVERSELANDSCAPE;
                }
            }
        }
    }
    catch (e)
    {
        Debug.write(e.message + "updateResolutionDevmodeValue() Error!");
    }
}


function convertAllFeaturesPrintTicketToDevmode(printTicket, scriptContext, devModeProperties, property)
{
    try
    {
        var xpathTable = getXPathTableObject( scriptContext );

        if( property && xpathTable )
        {
            var printTicketParser = new PrintTicketParser( printTicket, xpathTable );
            var features = property.getFeatureNameList().concat(CHANGFEATURE_TABLE_LIST);
            for(var i=0; i<features.length; i++ )
            {
                var xpath = getXPathString(features[i], xpathTable);
                if( xpath !== null )
                {
                    if( isArrayObject(xpath) )
                    {
                        for(var j=0; j<xpath.length; j++)
                        {
                            var node = printTicket.XmlNode.documentElement.selectSingleNode( xpath[j] );
                            if( node === null )
                            {
                                node = printTicket.XmlNode.documentElement.selectSingleNode( xpath[j].replace(/psf:Option\/@name/g,"psf:Option\/@psf:name") );
                            }
                            if( node !== null )
                            {
                                property.setCurrentValue(
                                        features[i],
                                        node.text,
                                        xpath[j].substring(xpath[j].indexOf("'")+1, xpath[j].indexOf("'", xpath[j].indexOf("'")+1) )
                                        );
                            }
                        }
                    }
                    else
                    {
                        if ("ns0000:OverlayWatermarkItem" == features[i]) {
                          printTicket.XmlNode.preserveWhiteSpace = 1;
                        }
                        var node = printTicket.XmlNode.documentElement.selectSingleNode( xpath );
                        if( containsInArray( features[i], AUTO_CREATE_FEATURE_LIST ) ||
                            containsInArray( features[i], AUTO_CREATE_PARAMETERDEF_LIST ) )
                        {
                            var fixvalue = property.getFixValue(features[i]);
                            if( fixvalue )
                            {
                                if( node !== null )
                                {
                                    node.text = fixvalue;
                                }
                            }
                            else
                            {
                                if( (features[i] == "psk:PageOrientation") &&
                                    (node !== null) )
                                {
                                    property.setLandscapeDirection( node.text );
                                }
                            }
                        }
                        else
                        {
                            if( node !== null )
                            {
                                property.setCurrentValue( features[i], node.text );
                            }
                            else
                            {
                                property.setCurrentValue( features[i], printTicketParser.getValue( features[i] ) );
                            }
                        }

                        if ("ns0000:OverlayWatermarkItem" == features[i]) {
                          printTicket.XmlNode.preserveWhiteSpace = 0;
                        }
                    }
                }
            }

            var newsetting = convertJsonObjectToString( property.devmode );

            devModeProperties.SetString("DocumentSettings", newsetting);
        }
    }
    catch(e)
    {

    }
    return;
}

function validateDevmodeToPrintTicket( printTicket, scriptContext )
{
    try
    {
        var xpathTable = getXPathTableObject( scriptContext );

        var printTicketParser = new PrintTicketParser( printTicket, xpathTable );

        validatePaperSizePrintTicket( printTicket );

        validateInputBinPrintTicket( printTicket );

        validateDuplexPrintTicket( printTicket );

        validateOutputColorPrintTicket( printTicket, printTicketParser );

        validateCustomPaperSizePrintTicket( printTicket, printTicketParser );
    }
    catch(e)
    {
        Debug.write(e.message + "validateDevmodeToPrintTicket() Error!");
    }
}

function validatePrintTicketToDevmode( printTicket, scriptContext )
{
    try
    {
        var xpathTable = getXPathTableObject( scriptContext );

        var printTicketParser = new PrintTicketParser( printTicket, xpathTable );

        validatePaperSizePrintTicket( printTicket );

        validateInputBinPrintTicket( printTicket );

        validateDuplexPrintTicket( printTicket );

        validateResolutionPrintTicket(printTicket, scriptContext);

        validateOutputColorPrintTicket( printTicket, printTicketParser );

        validateCustomPaperSizePrintTicket( printTicket, printTicketParser );
    }
    catch(e)
    {
        Debug.write(e.message + "validatePrintTicketToDevmode() Error!");
    }
}

function validatePrintCapabilities( printCapabilities )
{
    validateInputBinCapabilities( printCapabilities );

    validatePaperSizeCapabilities( printCapabilities );
}

function validatePaperSizePrintTicket( ptpc )
{
    try
    {
        var xpath = null;
        
        if( global.eventtype === EVENT_DEVMODETOPRINTTICKET ||
            global.eventtype === EVENT_PRINTCAPABILITIES )
        {
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='psk:OtherMetricFolio']";
            var othermetric = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (othermetric)
            {
                othermetric.attributes[0].text = "ns0000:FOLIO";
            }
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='psk:NorthAmericaCSheet']";
            othermetric = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (othermetric)
            {
                othermetric.attributes[0].text = "ns0000:Ansic";
            }
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='psk:NorthAmericaDSheet']";
            othermetric = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (othermetric)
            {
                othermetric.attributes[0].text = "ns0000:_22X34";
            }
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='psk:NorthAmericaESheet']";
            othermetric = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (othermetric)
            {
                othermetric.attributes[0].text = "ns0000:_34X44";
            }
        }
        else if( global.eventtype === EVENT_PRINTTICKETTODEVMODE )
        {
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='ns0000:FOLIO']";
            var folio = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (folio)
            {
                folio.attributes[0].text = "psk:OtherMetricFolio";
            }
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='ns0000:Ansic']";
            var csheet = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (csheet)
            {
                csheet.attributes[0].text = "psk:NorthAmericaCSheet";
            }
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='ns0000:_22X34']";
            var dsheet = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (dsheet)
            {
                dsheet.attributes[0].text = "psk:NorthAmericaDSheet";
            }
            xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Option[@name='ns0000:_34X44']";
            var esheet = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (esheet)
            {
                esheet.attributes[0].text = "psk:NorthAmericaESheet";
            }
        }
    }
    catch(e)
    {
        Debug.write(e.message + "validatePaperSizePrintTicket() Error!");
    }
}

function validateCustomPaperSizePrintTicket( printTicket, printTicketParser )
{
    var documentSize = printTicketParser.getValue( "psk:PageMediaSize" );
    var documentSizeCustomWidth  = printTicketParser.getValue( "psk:PageMediaSizeMediaSizeWidth" );
    var documentSizeCustomHeight = printTicketParser.getValue( "psk:PageMediaSizeMediaSizeHeight" );

    var printSizeCustomWidth  = printTicketParser.getValue( "ns0000:PageMediaSizeMediaSizeWidth" );
    var printSizeCustomHeight = printTicketParser.getValue( "ns0000:PageMediaSizeMediaSizeHeight" );

    if( documentSize === "psk:CustomMediaSize" &&
        documentSizeCustomWidth !== null && documentSizeCustomHeight !== null &&
        printSizeCustomWidth    !== null && printSizeCustomHeight    !== null )
    {
        var printWidthNode  = printTicketParser.getNode( "ns0000:PageMediaSizeMediaSizeWidth" );
        var printHeightNode = printTicketParser.getNode( "ns0000:PageMediaSizeMediaSizeHeight" );

        printWidthNode.text  = documentSizeCustomWidth;
        printHeightNode.text = documentSizeCustomHeight;
    }
}


function validateInputBinPrintTicket( ptpc )
{
    try
    {
        var xpath = null;
        
        if( global.eventtype === EVENT_DEVMODETOPRINTTICKET ||
            global.eventtype === EVENT_PRINTCAPABILITIES )
        {
            xpath = "psf:Feature[@name='psk:PageInputBin']/psf:Option[@name='psk:Manual']";
            var manual = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (manual)
            {
                manual.attributes[0].text = "ns0000:BypassTray";
                var removeNodes = manual.selectNodes("psf:ScoredProperty");
                if( removeNodes )
                {
                    for ( var i = 0; i < removeNodes.length; i++)
                    {
                        manual.removeChild( removeNodes[i] );
                    }
                }
            }
        }
        else if( global.eventtype === EVENT_PRINTTICKETTODEVMODE )
        {
            xpath = "psf:Feature[@name='psk:JobInputBin']";
            var inputbin = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (inputbin)
            {
                inputbin.attributes[0].text = "psk:PageInputBin";
            }

            xpath = "psf:Feature[@name='psk:PageInputBin']/psf:Option";
            var option = ptpc.XmlNode.documentElement.selectSingleNode( xpath );
            if (option)
            {
                if( option.getAttribute("name") == "psk:AutoSelect" )
                {
                    option.attributes[0].text = "ns0000:AUTO";
                }
                else if( option.getAttribute("name") == "ns0000:BypassTray" )
                {
                    option.attributes[0].text = "psk:Manual";
                }
            }
        }
    }
    catch(e)
    {
        Debug.write(e.message + "validateInputBinPrintTicket() Error!");
    }
}

function validateDuplexPrintTicket( printTicket )
{
    var xpath = null;

    xpath = "psf:Feature[@name='psk:JobDuplexAllDocumentsContiguously']/psf:Option";
    var contiguously = printTicket.XmlNode.documentElement.selectSingleNode(xpath);

    xpath = "psf:Feature[@name='ns0000:JobDuplexAllDocumentsDirection']/psf:Option";
    var direction = printTicket.XmlNode.documentElement.selectSingleNode(xpath);
    if (!direction)
    {
        var directionnode = printTicket.XmlNode.documentElement.selectNodes("psf:Feature");
        for (var i = 0; i < directionnode.length; i++)
        {
            if (directionnode[i].attributes[0].text == "ns0000:JobDuplexAllDocumentsDirection")
            {
                direction = directionnode[i].selectSingleNode("psf:Option");
                break;
            }
        }
    }

    if ((contiguously) && (direction))
    {
        if (contiguously.attributes[0].text != PSKJOBDUPLEX_CONTIGUOUSLY_MAP[direction.attributes[0].text])
        {
            direction.attributes[0].text = RPSKJOBDUPLEX_DIRECTION_MAP[contiguously.attributes[0].text][direction.attributes[0].text];
        }
    }
}

/**
*
*  Base64 encode / decode
*  http://www.webtoolkit.info/
*
**/
 

var Base64 = {

  // private property
  _keyStr: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",

  // public method for encoding
  encode: function (input) {
    var output = "";
    var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
    var i = 0;

    input = Base64._utf8_encode(input);

    while (i < input.length) {

      chr1 = input.charCodeAt(i++);
      chr2 = input.charCodeAt(i++);
      chr3 = input.charCodeAt(i++);

      enc1 = chr1 >> 2;
      enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
      enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
      enc4 = chr3 & 63;

      if (isNaN(chr2)) {
        enc3 = enc4 = 64;
      } else if (isNaN(chr3)) {
        enc4 = 64;
      }

      output = output +
      this._keyStr.charAt(enc1) + this._keyStr.charAt(enc2) +
      this._keyStr.charAt(enc3) + this._keyStr.charAt(enc4);

    }

    return output;
  },

  // public method for decoding
  decode: function (input) {
    var output = "";
    var chr1, chr2, chr3;
    var enc1, enc2, enc3, enc4;
    var i = 0;

    input = input.replace(/[^A-Za-z0-9\+\/\=]/g, "");

    while (i < input.length) {

      enc1 = this._keyStr.indexOf(input.charAt(i++));
      enc2 = this._keyStr.indexOf(input.charAt(i++));
      enc3 = this._keyStr.indexOf(input.charAt(i++));
      enc4 = this._keyStr.indexOf(input.charAt(i++));

      chr1 = (enc1 << 2) | (enc2 >> 4);
      chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
      chr3 = ((enc3 & 3) << 6) | enc4;

      output = output + String.fromCharCode(chr1);

      if (enc3 != 64) {
        output = output + String.fromCharCode(chr2);
      }
      if (enc4 != 64) {
        output = output + String.fromCharCode(chr3);
      }

    }

    output = Base64._utf8_decode(output);

    return output;

  },

  // private method for UTF-8 encoding
  _utf8_encode: function (string) {
    string = string.replace(/\r\n/g, "\n");
    var utftext = "";

    for (var n = 0; n < string.length; n++) {

      var c = string.charCodeAt(n);

      if (c < 128) {
        utftext += String.fromCharCode(c);
      }
      else if ((c > 127) && (c < 2048)) {
        utftext += String.fromCharCode((c >> 6) | 192);
        utftext += String.fromCharCode((c & 63) | 128);
      }
      else {
        utftext += String.fromCharCode((c >> 12) | 224);
        utftext += String.fromCharCode(((c >> 6) & 63) | 128);
        utftext += String.fromCharCode((c & 63) | 128);
      }

    }

    return utftext;
  },

  // private method for UTF-8 decoding
  _utf8_decode: function (utftext) {
    var string = "";
    var i = 0;
    var c = c1 = c2 = 0;

    while (i < utftext.length) {

      c = utftext.charCodeAt(i);

      if (c < 128) {
        string += String.fromCharCode(c);
        i++;
      }
      else if ((c > 191) && (c < 224)) {
        c2 = utftext.charCodeAt(i + 1);
        string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));
        i += 2;
      }
      else {
        c2 = utftext.charCodeAt(i + 1);
        c3 = utftext.charCodeAt(i + 2);
        string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
        i += 3;
      }

    }

    return string;
  }

}

function validateResolutionPrintTicket(printTicket, scriptContext) {
    var xpath = null;

    try
    {
        xpath = "psf:Feature[@name='psk:PageResolution']/psf:Option";
        var resolution = printTicket.XmlNode.documentElement.selectSingleNode(xpath);

        if ( (resolution) && (resolution.childNodes.length != 0) )
        {
          var pdl = scriptContext.DriverProperties.GetString("pdl");
          if (pdl === 'PC-FAX')
          {
            var availableResolutionMap = null;
            var str = scriptContext.QueueProperties.GetString("DeviceSettings");
            if( str )
            {
              var deviceSettings = Base64.decode(str);
              availableResolutionMap = ConvertToJsonObjectFromString(deviceSettings);
            }
            else
            {
              availableResolutionMap = FAX_DEFAULT_AVAILABLE_RESOLUTION_MAP;
            }
            var available = availableResolutionMap[FAX_RESOLUTION_MAP[resolution.attributes[0].text]];
            if( available !== 'on')
            {
              resolution.attributes[0].text = DEFAULT_FAX_RESOLUTION;
              resolution.childNodes[0].text = RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionX'];
              resolution.childNodes[1].text = RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionY'];
            }
          }
          else if (pdl === 'PCLXL')
          {
            var modelName = scriptContext.QueueProperties.GetString("PrinterModelName");
            if (modelName === 'Generic Model For Primax')
            {
              resolution.attributes[0].text = DEFAULT_PRIMAX_RESOLUTION;
              resolution.childNodes[0].text = RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionX'];
              resolution.childNodes[1].text = RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionY'];
            }
          }

          if ((RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionX'] != resolution.childNodes[0].text) && (RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionY'] != resolution.childNodes[1].text))
          {
            resolution.childNodes[0].text = RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionX'];
            resolution.childNodes[1].text = RESOLUTION_MAP[resolution.attributes[0].text]['ResolutionY'];
          }
        }
    }
    catch(e)
    {
        Debug.write(e.message + "validateResolutionPrintTicket() Error!");
    }
} 


function validateOutputColorPrintTicket(printTicket, printTicketParser) {
    try
    {
        var publicColor  = printTicketParser.getValue( "psk:PageOutputColor" );
        var privateColor = printTicketParser.getValue( "ns0000:PageDriverOutputColor" );
        var fixdmColor   = printTicketParser.getValue( "ns0000:JobFixdmColor" );

        if( publicColor && privateColor )
        {
            if( ! fixdmColor ||
                fixdmColor === "ns0000:Off" )
            {
                xpath = "psf:Feature[@name='ns0000:PageDriverOutputColor']";
                var outputcolor = printTicket.XmlNode.documentElement.selectSingleNode(xpath);
                if( !outputcolor )
                {
                    var outputcolor_node = printTicket.XmlNode.documentElement.selectNodes( "psf:Feature" );
                    for ( var i = 0; i < outputcolor_node.length; i++) {
                       if ( outputcolor_node[i].attributes[0].text == "ns0000:PageDriverOutputColor" ) {
                            outputcolor = outputcolor_node[i];
                            break;
                        }
                    }
                }

                if( outputcolor ) {
                    printTicket.XmlNode.documentElement.removeChild( outputcolor );
                }
            }
            else if( fixdmColor === "ns0000:On" )
            {
                var psknode = printTicketParser.getNode( "psk:PageOutputColor" );
                psknode.text = "psk:Color";
            }
        }
    }
    catch(e)
    {
        Debug.write(e.message + "validateOutputColorPrintTicket() Error!");
    }
}

function validateInputBinCapabilities(printCapabilities)
{
    validateInputBinPrintTicket( printCapabilities );
}

function validatePaperSizeCapabilities(printCapabilities)
{
    validatePaperSizePrintTicket( printCapabilities );
}

function removePrefix( withPrefix )
{
    return withPrefix.substring( withPrefix.indexOf(":") + 1 );
}

function createPropertyObject( scriptContext, devModeProperties, ptpc )
{
    var property = null;
    try
    {
        if( !scriptContext ) {
            return null;
        }

        var devmodeObject = null;
        if( devModeProperties ) {
            devmodeObject = getDevmodeObject( devModeProperties );
            if( !devmodeObject ) {
                devmodeObject = {};
            }
        }
 
        var modelName = GetCurrentModelName(scriptContext);
        var modelSpecObject = getModelSpecObject( scriptContext, modelName );

        var printSizeName = devmodeObject["ns0000:PagePrintPaperSize"];
        if( ( printSizeName ) && ( printSizeName.match( "ns0000:UserForm" ) ) )
        {
            modelSpecObject.modelSpec["ns0000:PagePrintPaperSize"].pickone.push(printSizeName);
        }

        var uiCustom = scriptContext.QueueProperties.GetString("UI_Customize");
        var rcf = null;
        var ptpsdMap = null;
        if( uiCustom.length > 0 )
        {
            var rcfDom = CreateDomDoc( ptpc );
            if( rcfDom.loadXML(uiCustom) )
            {
                rcf = new RCF( rcfDom );
                ptpsdMap = getPtPsdMapObject( scriptContext );
            }
        }

        property = new Property(    modelSpecObject,
                                        devmodeObject,
                                        rcf,
                                        ptpsdMap,
                                        (devModeProperties) ? true : false );

        property.validate();
    }
    catch(e)
    {
        property = null;
        Debug.write(e.message + "!!! createPropertyObject() Failed !!!");
    }

    return property;
}

function convertJsonObjectToString( json )
{
    var jsonstr = "";
    if( json && json != {} )
    {
        jsonstr = "{";
        for(var key in json )
        {
            if( isHashObject(json[key]) )
            {
                jsonstr += "\"" + key + "\":" + "{";
                for(var subkey in json[key] )
                {
                    var escstr = json[key][subkey].replace(/\\/g, '\\\\').replace(/\//g, '\\/').replace(/\"/g, '\\"');
                    jsonstr += "\"" + subkey + "\":\"" + escstr + "\",";
                }
                jsonstr = jsonstr.substr(0, jsonstr.length-1 );
                jsonstr += "},";
            }
            else
            {
                var escstr = json[key].replace(/\\/g, '\\\\').replace(/\//g, '\\/').replace(/\"/g, '\\"');
                jsonstr += "\"" + key + "\":\"" + escstr + "\",";
            }
        }
        jsonstr = jsonstr.substr(0, jsonstr.length-1 );
        jsonstr += "}";
    }

    return jsonstr;
}

function getXPathFromPrintSchemaTable(featureName, ps_table)
{
    var ret = null;
    for(var i=0; i<ps_table.length; i++)
    {
        if( featureName == ps_table[i][KEY_ATTRIBUTE] )
        {
            if( "xpath" in ps_table[i] ) {
                ret = ps_table[i]["xpath"];
            }
            break;
        }
    }
    return ret;
}

function getXPathString(feature, xpathTable)
{
    var ret = null;
    if (feature in xpathTable)
    {
        if (xpathTable[feature] === 0)
        {
            ret = null;
        }
        else if (xpathTable[feature] == 1)
        {
            ret = "psf:Feature[@name='" + feature + "']/psf:Option/@name";
        }
        else if (xpathTable[feature] == 2)
        {
            ret = "psf:ParameterInit[@name='" + feature + "']/psf:Value";
        }
        else if ( "3" in xpathTable[feature] )
        {
            for(var i=0; i<xpathTable[feature]["3"].length; i++ )
            {
                if (ret === null)
                {
                    ret = "psf:Feature[@name='" + xpathTable[feature]["3"][i] + "']/";
                }
                else
                {
                    ret += "psf:Feature[@name='" + xpathTable[feature]["3"][i] + "']/";
                }
            }
            ret += "psf:Feature[@name='" + feature + "']/psf:Option/@name";
        }
        else if ( "4" in xpathTable[feature] )
        {
            for(var i=0; i<xpathTable[feature]["4"].length; i++ )
            {
                if (ret === null)
                {
                    ret = "psf:Feature[@name='" + xpathTable[feature]["4"][i] + "']/";
                }
                else
                {
                    ret += "psf:Feature[@name='" + xpathTable[feature]["4"][i] + "']/";
                }
            }
            ret += "psf:Option/psf:ScoredProperty[@name='" + feature + "']/psf:Value";
        }
        else if ( "5" in xpathTable[feature] )
        {
            if( xpathTable[feature]['5'].length > 0 )
            {
                ret = [];
                for(var j=0; j<xpathTable[feature]['5'].length; j++)
                {
                    ret.push("psf:Feature[@name='" + xpathTable[feature]['5'][j] + "']/psf:Option/psf:ScoredProperty[@name='" + feature + "']/psf:Value");
                }
            }
        }
        else if ( "6" in xpathTable[feature] )
        {
            if( xpathTable[feature]['6'].length > 0 )
            {
                ret = [];
                for(var j=0; j<xpathTable[feature]['6'].length; j++)
                {
                    ret.push("psf:Feature[@name='" + xpathTable[feature]['6'][j] + "']/psf:Feature[@name='" + feature + "']/psf:Option/@name");
                }
            }
        }
        else
        {
        }
    }
    else
    {
        ret = null;
    }
    
    return ret;
}

function WideCharBytesToMultiCharByte( wideCharBytes )
{
    var multi = [];
    for(var i=0, j=0; i<wideCharBytes.length; i=i+2, j++)
    {
        multi[j] = (wideCharBytes[i + 1] << 8) + wideCharBytes[i];
    }
    return multi;
}

function ConvertToUTF8Bytes(bytes) {
    var byteUTF8 = [];
    for (var i = 0, j = 0; i < bytes.length; i++, j++) {
        if (bytes[i] <= 0x7f) {
            byteUTF8[j] = bytes[i];
        } else if (bytes[i] <= 0xdf) {
            byteUTF8[j] = ((bytes[i] & 0x1f) << 6);
            byteUTF8[j] += bytes[i + 1] & 0x3f;
            i++;
        } else if (bytes[i] <= 0xe0) {
            byteUTF8[j] = ((bytes[i + 1] & 0x1f) << 6) | 0x0800;
            byteUTF8[j] += bytes[i + 2] & 0x3f;
            i += 2;
        } else {
            byteUTF8[j] = ((bytes[i] & 0x0f) << 12);
            byteUTF8[j] += (bytes[i + 1] & 0x3f) << 6;
            byteUTF8[j] += bytes[i + 2] & 0x3f;
            i += 2;
        }
    }
    return byteUTF8;
}

function PtToPsdFeatureKey( ptKey, ptpsdTable )
{
    var psdKey = [];
    try
    {
        if( PTPSD_KEY_ATTRIBUTEMULTI in ptpsdTable )
        {
            for( var k=0; k<ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI].length; k++)
            {
                if( ptKey == ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][k][PTPSD_KEY_PTNAME] )
                {
                    for( var l=0; l<ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][k][PTPSD_KEY_PSDNAME].length; l++)
                    {
                        psdKey.push(ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][k][PTPSD_KEY_PSDNAME][l]);
                    }
                    return psdKey;
                }
            }
        }
     
        if( PTPSD_KEY_ATTRIBUTE in ptpsdTable )
        {
            for( var j=0; j<ptpsdTable[PTPSD_KEY_ATTRIBUTE].length; j++)
            {
                if( ptKey == ptpsdTable[PTPSD_KEY_ATTRIBUTE][j][PTPSD_KEY_PTNAME] )
                {
                    psdKey.push(ptpsdTable[PTPSD_KEY_ATTRIBUTE][j][PTPSD_KEY_PSDNAME]);
                    return psdKey;
                }
            }
        }
    }
    catch(e)
    {
        psdKey = null;
    }

    return null;
}

function PsdToPtValueKey( ptFeature, psdValue, ptpsdTable )
{
    var result = [];
    try
    {
        if( PTPSD_KEY_ATTRIBUTEMULTI in ptpsdTable )
        {
            for( var i=0; i<ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI].length; i++)
            {
                if( ptFeature == ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PTNAME] )
                {
                    if( "pickone" in ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i] )
                    {
                        for( var j=0; j<ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][RCF_ITEM_KEY_PICKONE].length; j++ )
                        {
                            if(psdValue[0] !== null)
                            {
                                if( ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME][0] == psdValue[0] )
                                {
                                    if( ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME][1] == psdValue[1] ||
                                        ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME][1] == "@ANY" ||
                                        ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME][1] == "@NOTSUPPORTED" ||
                                        psdValue[1] == null )
                                    {
                                        result.push( ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PTNAME] );
                                    }
                                }
                            }
                            else
                            {
                                if( ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME][1] == psdValue[1] ||
                                    ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME][1] == "@ANY" )
                                {
                                    result.push( ptpsdTable[PTPSD_KEY_ATTRIBUTEMULTI][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PTNAME] );
                                }
                            }
                        }
                    }
                    else
                    {
                        result.push(psdValue[0]);
                    }
                }
            }
        }

        if( result.length > 0 )
        {
            return result;
        }

        if( PTPSD_KEY_ATTRIBUTE in ptpsdTable )
        {
            for( var i=0; i<ptpsdTable[PTPSD_KEY_ATTRIBUTE].length; i++)
            {
                if( ptFeature == ptpsdTable[PTPSD_KEY_ATTRIBUTE][i][PTPSD_KEY_PTNAME] )
                {
                    if( "pickone" in ptpsdTable[PTPSD_KEY_ATTRIBUTE][i] )
                    {
                        for( var j=0; j<ptpsdTable[PTPSD_KEY_ATTRIBUTE][i][RCF_ITEM_KEY_PICKONE].length; j++ )
                        {
                            if( ptpsdTable[PTPSD_KEY_ATTRIBUTE][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PSDNAME] == psdValue[0] )
                            {
                                result.push( ptpsdTable[PTPSD_KEY_ATTRIBUTE][i][PTPSD_KEY_PICKONE][j][PTPSD_KEY_PTNAME] );
                                break;
                            }
                        }
                    }
                    else
                    {
                        ( ptpsdTable[PTPSD_KEY_ATTRIBUTE][i][PTPSD_KEY_PSDNAME].match(/_mm$/) )
                            ? result.push( psdValue[0] * 100 ) : result.push( psdValue[0] );
                    }
                    break;
                }
            }
        }
    }
    catch(e)
    {
        result = [];
    }
    return result;
}

function RCFParser_GetVersion(rcf)
{
    var version = null;
    try
    {
        var root = rcf.selectSingleNode("rcf");
        version = root.getAttribute("version");
    }
    catch(e)
    {
    }
    return version;
}

function RCFParser_GetCustomer(rcf)
{
    var customer = { "id": "" };
    try
    {
        var customerNode = rcf.selectSingleNode("rcf/customer");
        var id = customerNode.getAttribute("id");
        customer['id'] = id;
    }
    catch (e) {
        customer = null;
    }
    return customer;
}

function RCFParser_GetDeviceSettings(rcf)
{
    var devicesettings = { "drivername": "", "independent":"no", RCF_ITEM_KEY_ITEMLIST: [] };

    try {
        var devicesettingsNode = rcf.selectSingleNode("rcf/devicesettings");
        devicesettings['drivername'] = devicesettingsNode.getAttribute("drivername");
        devicesettings['independent'] = devicesettingsNode.getAttribute("independent");

        var itemNodes = rcf.selectNodes("rcf/devicesettings/item");
        var ary = [];
        for( var i=0; i < itemNodes.length; i++)
        {
            var item = { RCF_ITEM_KEY_NAME: "", RCF_ITEM_KEY_VALUE: "" };
            item[RCF_ITEM_KEY_NAME] = itemNodes[i].getAttribute("name");
            item[RCF_ITEM_KEY_VALUE]= itemNodes[i].getAttribute("value");
            ary[i] = item;
        }
        devicesettings[RCF_ITEM_KEY_ITEMLIST] = ary;

    }
    catch (e) {
    }
    return devicesettings;
}

function RCFParser_GetFeatureLock(rcf)
{
    var lockItems = [];
    try {
        var featureLockNode = rcf.selectSingleNode("rcf/featurelock");

        var itemNodes = rcf.selectNodes("rcf/featurelock/item");
        for (var i = 0; i < itemNodes.length; i++) {

            var item = { RCF_ITEM_KEY_NAME: "", RCF_ITEM_KEY_FIXVALUE: null, RCF_ITEM_KEY_PICKONE: null };
            item[RCF_ITEM_KEY_NAME] = itemNodes[i].getAttribute("name");
            item[RCF_ITEM_KEY_FIXVALUE] = itemNodes[i].getAttribute("fixvalue");
            if (item[RCF_ITEM_KEY_FIXVALUE] === null)
            {
                var pickoneNodes = itemNodes[i].selectNodes("pickone");
                var pickone = [];
                for (var j = 0; j < pickoneNodes.length; j++)
                {
                    pickone[j] = pickoneNodes[j].getAttribute("name");
                }
                item[RCF_ITEM_KEY_PICKONE] = pickone;
            }
            lockItems[i] = item;
        }
        
    }
    catch (e) {
    }
    return lockItems;
}

function RCFParser_GetModifyItem(rcf)
{
    var modifyItems = [];
    try {
        var itemNodes = rcf.selectNodes("rcf/modifyitem/item");

        for (var i = 0; i < itemNodes.length; i++) {

            var item = {};
            var name = itemNodes[i].getAttribute("name");
            var max  = null;
            var node  = itemNodes[i].selectSingleNode("string");
            if( node )
            {
                max  = node.getAttribute("max");
            }
            if( max && name)
            {
                item[RCF_ITEM_KEY_MAX]  = max;
                item[RCF_ITEM_KEY_NAME] = name;
                modifyItems.push( item );
            }
        }
    }
    catch (e) {
    }
    return modifyItems;
}


function CreateDomDoc(ptpc)
{
    var domDoc = ptpc.XmlNode.cloneNode(false);
    if (domDoc.parseError.errorCode !== 0) {
       var myErr = workNode.parseError;
       Debug.write("You have error " + myErr.reason);
    } else {
       Debug.write(domDoc.xml);
    }
    return domDoc;
}

function CreateAllPrintCapabilities(printTicket, scriptContext, printCapabilities)
{
    try {
        var modelName = GetCurrentModelName(scriptContext);
        if( modelName !== null )
        {
            var commonPCList  = getInnerSpecObject( scriptContext );

            var modelSpecObject = getModelSpecObject( scriptContext, modelName );


            var uiCustom = scriptContext.QueueProperties.GetString("UI_Customize");
            var rcf = null;
            var ptpsdMap = null;
            if( uiCustom.length > 0 )
            {
                var rcfDom = CreateDomDoc(printCapabilities);
                if( rcfDom.loadXML(uiCustom) )
                {
                    rcf = new RCF( rcfDom );
                    ptpsdMap = getPtPsdMapObject( scriptContext );
                }
            }

            var property = new Property(modelSpecObject, null, rcf, ptpsdMap);

            var userformlist = getUserFormCustomSizeList( printCapabilities );

            var table = getPrintTicketTreeTable( commonPCList, property, scriptContext, userformlist );
            table = getPagePrintPaperSize(table, printCapabilities, property);
            for( var i=0; i<table.length; i++ )
            {
                table[i].createPrintCapabilities(
                            printCapabilities, 
                            printCapabilities.XmlNode.lastChild
                            );
            }

            AddPSPageOrientationCapabilities( printCapabilities, property );
        }
    }
    catch (e) {

        var result = 0;
        result = e.number;
        Debug.write(e.message + "completePrintCapabilities() Error!");        
    }
    
    return;
}

function getUserFormCustomSizeList(printCapabilities) {

    try
    {
        var userformtable = {};

        var xpath = "psf:Feature[@name='psk:PageMediaSize']";
        var pagemediasize = printCapabilities.XmlNode.documentElement.selectSingleNode(xpath);
        if (pagemediasize)
        {
            for (var i = 0; i < pagemediasize.childNodes.length; i++)
            {
                if (pagemediasize.childNodes[i].attributes[0].text.match("ns0000:UserForm"))
                {
                    var widthNode  = pagemediasize.childNodes[i].selectSingleNode( "psf:ScoredProperty[@name='psk:MediaSizeWidth']" );
                    var heightNode = pagemediasize.childNodes[i].selectSingleNode( "psf:ScoredProperty[@name='psk:MediaSizeHeight']" );
                    if( widthNode && heightNode )
                    {
                        userformtable[pagemediasize.childNodes[i].attributes[0].text] = {};
                        userformtable[pagemediasize.childNodes[i].attributes[0].text].width  = widthNode.text;
                        userformtable[pagemediasize.childNodes[i].attributes[0].text].height = heightNode.text;
                    }
                }
            }
        }
    }
    catch (e)
    {
        Debug.write(e.message + "getUserFormCustomSizeList() Error!");
    }
    return userformtable;
}
 
function getPagePrintPaperSize(table, printCapabilities, property) {

    try
    {
        var userForm = [];

        var xpath = "psf:Feature[@name='psk:PageMediaSize']";
        var pagemediasize = printCapabilities.XmlNode.documentElement.selectSingleNode(xpath);
        var string;
        if (pagemediasize)
        {
            var index;
            for (index = 0; index < table.length; index++)
            {
                if (table[index][KEY_NAME] == "ns0000:PagePrintPaperSize")
                {
                    break;
                }
            }

            for (var i = 0; i < pagemediasize.childNodes.length; i++)
            {
                if (pagemediasize.childNodes[i].attributes[0].text.match("ns0000:UserForm"))
                {
                    var option = new OptionElement(pagemediasize.childNodes[i].attributes[0].text, property);

                    var sizecount = pagemediasize.childNodes[i].childNodes.length - 1;
                    for (var j = 0; j < sizecount; j++)
                    {
                        var intelement = new ScoredPropertyIntegerElement(
                                                userpskPrefix + ":" + removePrefix(pagemediasize.childNodes[i].childNodes[j].attributes[0].text),
                                                pagemediasize.childNodes[i].childNodes[j].text,
                                                property,
                                                "ns0000:PagePrintPaperSize"
                                                );
                        option[KEY_CHILDREN].push(intelement);
                    }

                    table[index][KEY_CHILDREN].splice(((table[index][KEY_CHILDREN].length) - 1), 0, option);
                }
            }
        }
    }
    catch (e)
    {
        Debug.write(e.message + "getPagePrintPaperSize() Error!");
    }

    return table;
}

function AddPSPageOrientationCapabilities( printCapabilities, property )
{
    try
    {
        if( property.isSupportedElement(FEATURE_PAGE_ORIENTATION, OPTION_REVERSELANDSCAPE) )
        {
            var xpath = "psf:Feature[@name= '" + FEATURE_PAGE_ORIENTATION + "']";
            var node = printCapabilities.XmlNode.documentElement.selectSingleNode( xpath );
            if( node )
            {
                var option = CreateOption( printCapabilities, removePrefix(OPTION_REVERSELANDSCAPE), null, pskPrefix );
                if( option )
                {
                    node.appendChild( option );
                }
            }
        }
    }
    catch(e)
    {
        Debug.write(e.message + "AddPSPageOrientationCapabilities() Error!");
    }
}

var PrintTicketParser = function( printTicket, xpathTable )
{
    this.printTicket = printTicket;
    this.xpathTable  = xpathTable;
    this.valueList   = {};
    this.nodeList    = {};
    this.parse();
};

PrintTicketParser.prototype.parse = function() {
    var featureNodes = this.printTicket.XmlNode.documentElement.selectNodes("psf:Feature");
    try
    {
        for (var i = 0; i < featureNodes.length; i++)
        {
            var parentFeatureNameList = [];
            this.parseFeatureNode( featureNodes[i], parentFeatureNameList );
        }

        var parameterInitNodes = this.printTicket.XmlNode.documentElement.selectNodes("psf:ParameterInit");
        for (var i = 0; i < parameterInitNodes.length; i++)
        {
            var parameterName = parameterInitNodes[i].attributes[0].text;
            if (this.xpathTable[ parameterName ] == 2)
            {
                var value = parameterInitNodes[i].selectSingleNode("psf:Value");
                this.valueList[parameterName] = value.text;
                this.nodeList[parameterName]  = value;
            }
        }
    }
    catch(e)
    {
        Debug.write(e.message + "PrintTicketParser::parse() Error!");
    }
};

PrintTicketParser.prototype.parseFeatureNode = function( featureNode, parentFeatureNameList ) {

    try
    {
        var featureName = featureNode.attributes[0].text;

        var option = null;
        if( featureName in this.xpathTable )
        {
            if (this.xpathTable[ featureName ] == 1 && parentFeatureNameList.length == 0)
            {
                option = featureNode.selectSingleNode("psf:Option");
                this.valueList[featureName] = option.attributes[0].text;
                this.nodeList[featureName]  = option.attributes[0];
            }
            else if ( "3" in this.xpathTable[featureName] &&
                      parentFeatureNameList.length === this.xpathTable[featureName]['3'].length )
            {
                var match = true;
                for(var i=0; i<parentFeatureNameList.length; i++)
                {
                    if( parentFeatureNameList[i] !== this.xpathTable[featureName]['3'][i] )
                    {
                        match = false;
                        break;
                    }
                }

                if( match === true )
                {
                    option = featureNode.selectSingleNode("psf:Option");
                    this.valueList[featureName] = option.attributes[0].text;
                    this.nodeList[featureName]  = option.attributes[0];
                }
            }
        }

        var childFeatureNodes = featureNode.selectNodes("psf:Feature");
        parentFeatureNameList.push( featureName );
        for (var j = 0; j < childFeatureNodes.length; j++)
        {
            var parentSubFeatureNameList = parentFeatureNameList.concat();
            this.parseFeatureNode( childFeatureNodes[j], parentSubFeatureNameList );
        }
    }
    catch(e)
    {
        Debug.write(e.message + "PrintTicketParser::parseFeatureNode() Error : " + featureName);
    }
};

PrintTicketParser.prototype.getValue = function( name ) {
    if( name in this.valueList )
    {
        return this.valueList[name];
    }
    else
    {
        return null;
    }
};

PrintTicketParser.prototype.getNode = function( name ) {
    if( name in this.nodeList )
    {
        return this.nodeList[name];
    }
    else
    {
        return null;
    }
};

function containsInArray( value, arr )
{
    for(var i in arr)
    {
        if( arr.hasOwnProperty(i) && arr[i] === value){
            return true;
        }
    }
    return false;
}


function CreateElement( xmlNode, tag, tagNs )
{
    return xmlNode.createNode(1, tag, tagNs);
}

function CreateAttribute(xmlNode, attributeValue)
{
    var attributeNode = null;
    if (attributeValue !== "")
    {
        attributeNode = xmlNode.createNode(2, GetAttributeName(attributeValue), GetAttributeNameSpace(attributeValue));
        if (attributeNode) {
            attributeNode.value = attributeValue;
        }
    }
    return attributeNode;
}

function GetAttributeName( attributeValue )
{
    if( attributeValue.indexOf("xsd") === 0 )
    {
        return "xsi:type";
    }
    else
    {
        return "name";
    }
}

function GetAttributeNameSpace(attributeValue) {
    if (attributeValue.indexOf("xsd") === 0) {
        return SCHEMA_INST_URI;
    }
    else {
        return FRAMEWORK_URI;
    }
}

function GetPrefixURI(prefix)
{
    if( prefix == "psk")
    {
        return  pskNs;
    }
    else
    {
        return ricohNs;
    }
    return null;
}



function GetCurrentModelName(scriptContext)
{
    var model = GENERIC_MODEL_NAME;
    try {
        var drivertype = GetDriverType(scriptContext);
        if( drivertype == "ud" )
        {
            model = GetStringFromProperties(scriptContext.QueueProperties, "PrinterModelName");
        }
        else if( drivertype == "generic" || drivertype == "model" )
        {
            model = GetStringFromProperties( scriptContext.QueueProperties, "Config:PrinterModelName" );
            if( model != null )
            {
                model = model.replace( /_/g, " " );
            }
            else
            {
                model = GENERIC_MODEL_NAME;
            }
        }
        else
        {
            model = GENERIC_MODEL_NAME;
        }
    }
    catch (e) {
        model = "Generic Model";
        var result = 0;
        result = e.number;
        Debug.write(e.message + "GetCurrentModelName() Error!");
    }

    return model;
}

function GetDriverType(scriptContext)
{
    var result = null;
    var drivertype = GetStringFromProperties( scriptContext.QueueProperties, "drivertype" );
    if( "ud" == drivertype )
    {
        result = "ud";
    }
    else
    {
        var config_drivertype = GetStringFromProperties( scriptContext.QueueProperties, "Config:drivertype" );
        result = (config_drivertype == null) ? drivertype : config_drivertype;
    }
    return result;
}

function GetModelSpecTable(scriptContext, modelName)
{

    var modelSpecObject = null;
    try
    {
        modelSpecObject = GetStringFromProperties( scriptContext.DriverProperties, "spec_" + modelName );
        if( modelSpecObject == null )
        {
            var clone_table = scriptContext.DriverProperties.GetString( "clone_model_table" );
            var clone_table_object = ConvertToJsonObjectFromString( clone_table );
            modelSpecObject = scriptContext.DriverProperties.GetString( "spec_" + clone_table_object[modelName] );
        }
    }
    catch (e)
    {
        modelSpecObject = scriptContext.DriverProperties.GetString( "spec_" + GENERIC_MODEL_NAME );
    }

    return modelSpecObject;
}

function GetModelSpecTableFromStream(scriptContext, modelName)
{
    var modelSpecObject = null;
    var modelSpecBytes = [];
    try
    {
        modelSpecBytes  = GetReadStreamFromProperties( scriptContext.DriverProperties, "spec_" + modelName + ".json" );
        modelSpecObject = BytesToString(modelSpecBytes);
        if( modelSpecObject == null )
        {
            var clone_table = scriptContext.DriverProperties.GetString( "clone_model_table" );
            var clone_table_object = ConvertToJsonObjectFromString( clone_table );

            modelSpecBytes  = GetReadStreamFromProperties( scriptContext.DriverProperties, "spec_" + clone_table_object[modelName] + ".json" );
            modelSpecObject = BytesToString(modelSpecBytes);
        }
    }
    catch (e)
    {
        modelSpecObject = scriptContext.DriverProperties.GetString( "spec_" + GENERIC_MODEL_NAME );
    }

    return modelSpecObject;
}



function GetBytesFromProperties( properties, name)
{
    var bytes = null;
    try
    {
        bytes = properties.GetBytes(name);
    }
    catch (e) {
        var result = 0;
        result = e.number;
        Debug.write(e.message + "GetBytesFromProperties() Error!");
    }
    return bytes;
}


function GetBoolFromProperties( properties, name)
{
    var value = null;
    try
    {
        value = properties.GetBool(name);
    }
    catch (e) {
        var result = 0;
        result = e.number;
        Debug.write(e.message + "GetBytesFromProperties() Error!");
    }
    return value;
}

function GetStringFromProperties( properties, name)
{
    var value = null;
    try
    {
        value = properties.GetString(name);
        if( value === "\r" )
        {
            value = "";
        }
    }
    catch (e) {
        value = null;
        Debug.write(e.message + "GetStringFromProperties() Error!");
    }
    return value;
}

function GetReadStreamFromProperties( properties, name)
{
    var value = [];
    try
    {
        var READ_STREAM_SIZE = 10000;
        var stream = properties.GetReadStream(name);
        do {
            //var start1 = new Date();
            var bytes = stream.Read(READ_STREAM_SIZE);
            //var end1 = new Date();
            //Debug.write("stream.Read() : " + (end1 - start1) + " ms\n");
            //var start2 = new Date();
            value = value.concat(bytes);
            //var end2 = new Date();
            //Debug.write("concat() : " + (end2 - start2) + " ms\n");
        } while ( bytes.length == READ_STREAM_SIZE );
    }
    catch (e) {
        value = [];
        Debug.write(e.message + "GetReadStreamFromProperties() Error!");
    }
    return value;
}

function ConvertToJsonObjectFromBytes(bytes)
{
    var byteString = String.fromCharCode.apply(String, bytes);
    return (new Function("return " + byteString))();
}

function ConvertToJsonObjectFromString(str)
{
    return (new Function("return " + str))();
}

function BytesToString(bytes)
{
    if( bytes && bytes.length > 0 )
    {
        return String.fromCharCode.apply(String, bytes);
    }
    else
    {
        return null;
    }
}

function isDefaultRequestEvent(scriptContext)
{
    var ret = false;
    try
    {
        var isChangeDefault = GetBoolFromProperties(scriptContext.QueueProperties, PROPERTY_KEY_ENABLE_CHANGEDEFAULT );
        var isModelChanged  = (GetStringFromProperties(scriptContext.QueueProperties, PROPERTY_KEY_MODE_MODELCHANGED ) === "changed") ? true : false;
        ret = (isChangeDefault || isModelChanged ) ? true : false;
    }
    catch(e)
    {
        ret = false;
    }
    return ret;
}

function createPageWatermark( scriptContext, property)
{
    Debug.write("createPageWatermark() is called!");

    var result = null;
    try
    {
        var wmlist = {};
        wmlist[KEY_TAG]       = TAG_FEATURE;
        wmlist[KEY_ATTRIBUTE] = "ns0000:PageWatermark";
        wmlist[KEY_PICKONE]   = {};

        var WMK_ITEM_FEATURE = "ns0000:WatermarkItem";
        var WMK_ITEM_TYPE    = TYPE_STRING;

        //
        var presets = getWatermarkList( scriptContext, property );
        //

        var scoredList = [];
        for(var i=0; i<presets.length; i++ )
        {
            scoredList[i] = {
                "name":WMK_ITEM_FEATURE,
                "type":WMK_ITEM_TYPE,
                "value":presets[i]
            };
        }
        wmlist[KEY_PICKONE]['ns0000:On'] = scoredList;

        var elementList = createFeatureElement( wmlist, property );

        if( elementList.length == 1 )
        {
            result = elementList[0];
        }
    }
    catch(e)
    {
        result = null;
    }
    return result;
}

function getWatermarkList( scriptContext, property )
{
    var watermarkpresets = {
        "english": ["CONFIDENTIAL","ORIGINAL","DRAFT","URGENT","COPY"],
        "japanese": ["CONFIDENTIAL","社外秘","DRAFT","マル秘","COPY"]
    };

    try
    {
        var presets = [];
        if( global.eventtype == EVENT_DEVMODETOPRINTTICKET )
        {
            var itemString = property.getCurrentValue( "ns0000:WatermarkItem" );
            if( itemString != null && itemString != "" )
            {
                presets.push( itemString );
            }
        }
        else if( global.eventtype == EVENT_PRINTCAPABILITIES )
        {
            if( global.language ) {
                presets = watermarkpresets[global.language];
            }
            if( !presets || presets.length == 0 )
            {
                presets = watermarkpresets["english"];
            }

            var settings = scriptContext.UserProperties.getString("watermark_settings");
            var settingsObj  = ConvertToJsonObjectFromString(settings);

            var itemList = settingsObj['personal_watermarks'];
            for(var i=0; i<itemList.length; i++)
            {
                var itemStr = scriptContext.UserProperties.getString(itemList[i]);
                var itemObj = ConvertToJsonObjectFromString(itemStr);
                presets.push( itemObj['settings']['name'] );
            }
        }
        else
        {
            // no
        }
    }
    catch(e)
    {
        Debug.write("Failed: getWatermarkPresetList");
    }

    return presets;
}

function createPageOverlayWatermark( scriptContext, property )
{
    Debug.write("createPageOverlayWatermark() is called!");

    var result = null;
    try
    {
        var supportlevel = property.getCurrentValue("ns0000:OverlayWatermarkSupportLevel");
        if (supportlevel == null || supportlevel != "1")
        {
            return createPageOverlayWatermarkEx(scriptContext, property);
        }

        var WMK_ITEM_BRINGTOFRONT = "ns0000:OverlayWatermarkBringToFront";
        var WMK_ITEM_SELECT = "ns0000:SelectOverlayWatermark";
        var WMK_ITEM_ALL_FEATURE = "ns0000:OverlayWatermarkItemAll";
        var WMK_ITEM_1ST_FEATURE = "ns0000:OverlayWatermarkItem1stPage";
        var WMK_ITEM_2ND_FEATURE = "ns0000:OverlayWatermarkItem2ndPage";
        var WMK_ITEM_3RD_FEATURE = "ns0000:OverlayWatermarkItem3rdPage";
        var WMK_ITEM_ENABLE_1ST = "ns0000:EnableOverlayWatermark1stPage";
        var WMK_ITEM_ENABLE_2ND = "ns0000:EnableOverlayWatermark2ndPage";
        var WMK_ITEM_ENABLE_3RD = "ns0000:EnableOverlayWatermark3rdPage";
        var WMK_VALUE_ALLPAGE = "ns0000:AllPagesSame";
        var WMK_VALUE_EACHPAGE = "ns0000:SelectForEachPage";

        var wmlist = {};
        wmlist[KEY_TAG]       = TAG_FEATURE;
        wmlist[KEY_ATTRIBUTE] = "ns0000:PageOverlayWatermark";
        wmlist[KEY_CHILDFEATURE] = [];

        var selectlist = {};
        selectlist[KEY_TAG]       = TAG_FEATURE;
        selectlist[KEY_ATTRIBUTE] = WMK_ITEM_SELECT;
        selectlist[KEY_PICKONE]   = {};

        var enable1stpage = {};
        enable1stpage[KEY_TAG]       = TAG_FEATURE;
        enable1stpage[KEY_ATTRIBUTE] = WMK_ITEM_ENABLE_1ST;
        enable1stpage[KEY_PICKONE]   = {};

        var enable2ndpage = {};
        enable2ndpage[KEY_TAG]       = TAG_FEATURE;
        enable2ndpage[KEY_ATTRIBUTE] = WMK_ITEM_ENABLE_2ND;
        enable2ndpage[KEY_PICKONE]   = {};

        var enable3rdpage = {};
        enable3rdpage[KEY_TAG]       = TAG_FEATURE;
        enable3rdpage[KEY_ATTRIBUTE] = WMK_ITEM_ENABLE_3RD;
        enable3rdpage[KEY_PICKONE]   = {};

        var bringtofront = {};
        bringtofront[KEY_TAG] = TAG_FEATURE;
        bringtofront[KEY_ATTRIBUTE] = WMK_ITEM_BRINGTOFRONT;
        bringtofront[KEY_PICKONE] = {};

        if (global.eventtype == EVENT_DEVMODETOPRINTTICKET)
        {
            var overlaypresets = {
                "english": ["(Add New)"],
                "japanese": ["（新規追加）"]
            };
            if (global.language) {
                default_namestring = overlaypresets[global.language];
            }
            if (!default_namestring || default_namestring.length == 0) {
                default_namestring = overlaypresets["english"];
            }

            var b2f = property.getCurrentValue(WMK_ITEM_BRINGTOFRONT);
            if (b2f != null && b2f != "")
            {
                bringtofront[KEY_PICKONE] = b2f;
            }

            selectlist[KEY_CHILDFEATURE] = [];

            var checkbox1 = property.getCurrentValue(WMK_ITEM_ENABLE_1ST);
            if (checkbox1 != "ns0000:On")
            {
                checkbox1 = "ns0000:Off";
            }
            var itemString1 = property.getCurrentValue(WMK_ITEM_1ST_FEATURE);
            if (null == itemString1 || itemString1.length == 0)
            {
                itemString1 = default_namestring;
            }
            enable1stpage[KEY_PICKONE][checkbox1] = [{ "name": WMK_ITEM_1ST_FEATURE, "type": TYPE_STRING, "value": itemString1 }];
            selectlist[KEY_CHILDFEATURE].push(enable1stpage);

            var checkbox2 = property.getCurrentValue(WMK_ITEM_ENABLE_2ND);
            if (checkbox2 != "ns0000:On")
            {
                checkbox2 = "ns0000:Off";
            }
            var itemString2 = property.getCurrentValue(WMK_ITEM_2ND_FEATURE);
            if (null == itemString2 || itemString2.length == 0)
            {
                itemString2 = default_namestring;
            }
            enable2ndpage[KEY_PICKONE][checkbox2] = [{ "name": WMK_ITEM_2ND_FEATURE, "type": TYPE_STRING, "value": itemString2 }];
            selectlist[KEY_CHILDFEATURE].push(enable2ndpage);

            var checkbox3 = property.getCurrentValue(WMK_ITEM_ENABLE_3RD);
            if (checkbox3 != "ns0000:On")
            {
                checkbox3 = "ns0000:Off";
            }
            var itemString3 = property.getCurrentValue(WMK_ITEM_3RD_FEATURE);
            if (null == itemString3 || itemString3.length == 0)
            {
                itemString3 = default_namestring;
            }
            enable3rdpage[KEY_PICKONE][checkbox3] = [{ "name": WMK_ITEM_3RD_FEATURE, "type": TYPE_STRING, "value": itemString3 }];
            selectlist[KEY_CHILDFEATURE].push(enable3rdpage);

            selectlist[KEY_PICKONE][WMK_VALUE_EACHPAGE] = {}

            var itemStringAll = property.getCurrentValue(WMK_ITEM_ALL_FEATURE);
            if (null == itemStringAll || itemStringAll.length == 0)
            {
                itemStringAll = default_namestring;
            }
            selectlist[KEY_PICKONE][WMK_VALUE_ALLPAGE] = [{ "name": WMK_ITEM_ALL_FEATURE, "type": TYPE_STRING, "value": itemStringAll }];
        }
        else
        {
            var presets = getOverlayWatermarkList(scriptContext, property);

            selectlist[KEY_PICKONE][WMK_VALUE_ALLPAGE] = createOverlayWatermarkScoredPropertyList(presets, WMK_ITEM_ALL_FEATURE);

            enable1stpage[KEY_PICKONE]['ns0000:On'] = createOverlayWatermarkScoredPropertyList(presets, WMK_ITEM_1ST_FEATURE);
            enable2ndpage[KEY_PICKONE]['ns0000:On'] = createOverlayWatermarkScoredPropertyList(presets, WMK_ITEM_2ND_FEATURE);
            enable3rdpage[KEY_PICKONE]['ns0000:On'] = createOverlayWatermarkScoredPropertyList(presets, WMK_ITEM_3RD_FEATURE);

            selectlist[KEY_PICKONE][WMK_VALUE_EACHPAGE] = {};
            selectlist[KEY_CHILDFEATURE] = [enable1stpage, enable2ndpage, enable3rdpage];
        }

        wmlist[KEY_CHILDFEATURE] = [bringtofront, selectlist];

        var elementList = createFeatureElement(wmlist, property);

        if( elementList.length == 1 )
        {
            result = elementList[0];
        }
    }
    catch(e)
    {
        result = null;
    }
    return result;
}

function createOverlayWatermarkScoredPropertyList( presets, featureName )
{
    var scoredList = [];
    for(var i=0; i<presets.length; i++ )
    {
        scoredList[i] = {
            "name":featureName,
            "type":TYPE_STRING,
            "value":presets[i]
        };
    }
    return scoredList;
}

function createPageOverlayWatermarkEx( scriptContext, property )
{
    Debug.write("createPageOverlayWatermark() is called!");

    var result = null;
    try
    {
        var wmlist = {};
        wmlist[KEY_TAG]       = TAG_FEATURE;
        wmlist[KEY_ATTRIBUTE] = "ns0000:PageOverlayWatermark";
        wmlist[KEY_PICKONE]   = {};

        var WMK_ITEM_FEATURE = "ns0000:OverlayWatermarkItem";
        var WMK_ITEM_TYPE    = TYPE_STRING;

        //
        var presets = getOverlayWatermarkList( scriptContext, property );
        //

        var scoredList = [];
        for(var i=0; i<presets.length; i++ )
        {
            scoredList[i] = {
                "name":WMK_ITEM_FEATURE,
                "type":WMK_ITEM_TYPE,
                "value":presets[i]
            };
        }
        wmlist[KEY_PICKONE]['ns0000:On'] = scoredList;

        var elementList = createFeatureElement( wmlist, property );

        if( elementList.length == 1 )
        {
            result = elementList[0];
        }
    }
    catch(e)
    {
        result = null;
    }
    return result;
}

function getOverlayWatermarkList( scriptContext, property )
{
    var watermarkpresets = {
        "english": ["New Overlay Data"],
        "japanese": ["新しい合成用データ"]
    };

    try
    {
        var presets = [];
        if( global.eventtype == EVENT_DEVMODETOPRINTTICKET )
        {
            var itemString = property.getCurrentValue( "ns0000:OverlayWatermarkItem" );
            if( itemString != null && itemString != "" )
            {
                presets.push( itemString );
            }
        }
        else if( global.eventtype == EVENT_PRINTCAPABILITIES )
        {
            if( global.language ) {
                presets = watermarkpresets[global.language];
            }
            if( !presets || presets.length == 0 )
            {
                presets = watermarkpresets["english"];
            }

            var settings = scriptContext.UserProperties.getString("overlaywatermark_settings");
            var settingsObj  = ConvertToJsonObjectFromString(settings);

            var itemList = settingsObj['personal_list'];
            for(var i=0; i<itemList.length; i++)
            {
                var itemStr = scriptContext.UserProperties.getString(itemList[i]);
                var itemObj = ConvertToJsonObjectFromString(itemStr);
                presets.push( itemObj['settings']['name'] );
            }
        }
        else
        {
            // no
        }
    }
    catch(e)
    {
        Debug.write("Failed: getOverlayWatermarkList");
    }

    return presets;
}

function createPageImageWatermark( scriptContext, property)
{
    Debug.write("createPageImageWatermark() is called!");

    var result = null;
    try
    {
        var wmlist = {};
        wmlist[KEY_TAG]       = TAG_FEATURE;
        wmlist[KEY_ATTRIBUTE] = "ns0000:PageImageWatermark";
        wmlist[KEY_PICKONE]   = {};

        var WMK_ITEM_FEATURE = "ns0000:ImageWatermarkItem";
        var WMK_ITEM_TYPE    = TYPE_STRING;

        //
        var presets = getImageWatermarkList( scriptContext, property );
        //

        var scoredList = [];
        for(var i=0; i<presets.length; i++ )
        {
            scoredList[i] = {
                "name":WMK_ITEM_FEATURE,
                "type":WMK_ITEM_TYPE,
                "value":presets[i]
            };
        }
        wmlist[RCF_ITEM_KEY_PICKONE]['ns0000:On'] = scoredList;

        var elementList = createFeatureElement( wmlist, property );

        if( elementList.length == 1 )
        {
            result = elementList[0];
        }
    }
    catch(e)
    {
        result = null;
    }
    return result;
}

function getImageWatermarkList( scriptContext, property )
{
    var watermarkpresets = {
        "english": ["New Bitmap"],
        "japanese": ["新しいイメージスタンプ"]
    };

    try
    {
        var presets = [];
        if( global.eventtype == EVENT_DEVMODETOPRINTTICKET )
        {
            var itemString = property.getCurrentValue( "ns0000:ImageWatermarkItem" );
            if( itemString != null && itemString != "" )
            {
                presets.push( itemString );
            }
        }
        else if( global.eventtype == EVENT_PRINTCAPABILITIES )
        {
            if( global.language ) {
                presets = watermarkpresets[global.language];
            }
            if( !presets || presets.length == 0 )
            {
                presets = watermarkpresets["english"];
            }

            var settings = scriptContext.UserProperties.getString("imagewatermark_settings");
            var settingsObj  = ConvertToJsonObjectFromString(settings);

            var itemList = settingsObj['personal_list'];
            for(var i=0; i<itemList.length; i++)
            {
                var itemStr = scriptContext.UserProperties.getString(itemList[i]);
                var itemObj = ConvertToJsonObjectFromString(itemStr);
                presets.push( itemObj['settings']['name'] );
            }
        }
        else
        {
            // no
        }
    }
    catch(e)
    {
        Debug.write("Failed: getOverlayWatermarkList");
    }

    return presets;
}

function createPageColorProfile( scriptContext, property )
{
    Debug.write("createPageColorProfile() is called!");

    var feature = "ns0000:PageColorProfile";

    return createColorProfileFeature( scriptContext, property, feature );
}

function createPageColorProfileForText( scriptContext, property )
{
    Debug.write("createPageColorProfileForText() is called!");

    var feature = "ns0000:PageColorProfileForText";

    return createColorProfileFeature( scriptContext, property, feature );
}

function createPageColorProfileForGraphics( scriptContext, property )
{
    Debug.write("createPageColorProfileForGraphics() is called!");

    var feature = "ns0000:PageColorProfileForGraphics";

    return createColorProfileFeature( scriptContext, property, feature );
}

function createPageColorProfileForPhoto( scriptContext, property )
{
    Debug.write("createPageColorProfileForPhoto() is called!");

    var feature = "ns0000:PageColorProfileForPhoto";

    return createColorProfileFeature( scriptContext, property, feature );
}

function createColorProfileFeature( scriptContext, property, featureName )
{
    //debugger;
    Debug.write("createColorProfileFeature() is called!");

    var result = null;
    try
    {
        if( property.isSupported(featureName) )
        {
            var defaultPresets = property.getPickoneList( featureName );

            var pslist = {};
            pslist[KEY_TAG]       = TAG_FEATURE;
            pslist[KEY_ATTRIBUTE] = featureName;
            pslist[KEY_PICKONE]   = {};

            var WMK_ITEM_FEATURE = "ns0000:ProfileName";
            var WMK_ITEM_TYPE    = TYPE_STRING;

            var scoredList = [];
            if( global.eventtype == EVENT_DEVMODETOPRINTTICKET )
            {
                //for debug
                //property.setCurrentValue( "ns0000:ProfileName", "automatic" );
                var itemString = property.getCurrentValue( featureName );
                if( itemString != null && itemString != "" )
                {
                    if( itemString === "ns0000:AddProfile" )
                    {
                        //TBD
                        property.modelSpec.modelSpec[featureName].pickone.push("ns0000:AddProfile");
                        scoredList[0] = {
                            "name":WMK_ITEM_FEATURE,
                            "type":WMK_ITEM_TYPE,
                            "value":property.getCurrentValue("ns0000:ProfileName", featureName)
                        };
                        pslist[KEY_PICKONE]['ns0000:AddProfile'] = scoredList;
                    }
                }
            }
            else if( global.eventtype == EVENT_PRINTCAPABILITIES )
            {
                //
                var presets = getColorProfileList( scriptContext, property );
                //

                if (null != presets && 0 < presets.length)
                {
                    for(var i=0; i<presets.length; i++ )
                    {
                        scoredList[i] = {
                            "name":WMK_ITEM_FEATURE,
                            "type":WMK_ITEM_TYPE,
                            "value":presets[i]
                        };
                    }
                    pslist[KEY_PICKONE]['ns0000:AddProfile'] = scoredList;
                    //TBD
                    property.modelSpec.modelSpec[featureName].pickone.push("ns0000:AddProfile");
                }
            }

            var elementList = createFeatureElement( pslist, property );

            if( elementList.length == 1 )
            {
                result = elementList[0];
            }
        }
    }
    catch(e)
    {
        result = null;
    }
    return result;
}

function getColorProfileList( scriptContext, property )
{
    var presets = [];
    try
    {
        var presets = [];
        if( global.eventtype == EVENT_DEVMODETOPRINTTICKET )
        {
            var itemString = property.getCurrentValue( "ns0000:ProfileName" );
            if( itemString != null && itemString != "" )
            {
                presets.push( itemString );
            }
        }
        else if( global.eventtype == EVENT_PRINTCAPABILITIES )
        {
            var settings = scriptContext.UserProperties.getString("colorprofile_settings");
            if (null != settings && "" != settings)
            {
                var settingsObj  = ConvertToJsonObjectFromString(settings);

                var itemList = settingsObj['personal_list'];
                for(var i=0; i<itemList.length; i++)
                {
                    var itemStr = scriptContext.UserProperties.getString(itemList[i]);
                    if (null != itemStr && "" != itemStr)
                    {
                        var itemObj = ConvertToJsonObjectFromString(itemStr);
                        presets.push( itemObj['settings']['managed_colorprofile'] );
                    }
                }
            }
        }
        else
        {
            // no
        }

    }
    catch(e)
    {
        Debug.write("Failed: getColorProfileList");
    }

    return presets;
}

function setCurrentLanguage( printCapabilities )
{
    var langList = { "用紙サイズ":"japanese" };
    var xpath = "psf:Feature[@name='psk:PageMediaSize']/psf:Property[@name='psk:DisplayName']";
    var node = printCapabilities.XmlNode.documentElement.selectSingleNode( xpath );

    global.language = langList[node.text];
}

function rewriteDisplayName( printCapabilities )
{
    try
    {
        var xpath = "psf:Feature[@name='psk:PageMediaType']/psf:Option/psf:Property[@name='psk:DisplayName']/psf:Value";
        var pagemediatype = printCapabilities.XmlNode.documentElement.selectNodes(xpath);
        if (pagemediatype)
        {
            for (var i = 0; i < pagemediatype.length; i++)
            {
                for (var j = 0; j < pagemediatype[i].childNodes.length; j++)
                {
                    if(pagemediatype[i].childNodes[j].text.indexOf("&&") != -1)
                    {
                        var result = pagemediatype[i].childNodes[j].text.replace( /&&/g , "&" ) ;
                        pagemediatype[i].childNodes[j].text = result;
                    }
                }
            }
        }
    }
    catch (e)
    {
        Debug.write(e.message + "rewriteDisplayName() Error!");    
    }
}

function setAdditionalFeatureString_Generic(jsSttring)
{
    var result;
    result = jsSttring.slice(0, -1) + "," + JOBEXTRAUISETTINGS_GENERIC  + "," + JOBENFORCEUSERCODE + "}";
    return result;
}

function setAdditionalFeatureString(jsSttring)
{
    var result;
    var index = jsSttring.lastIndexOf('}');
    result = jsSttring.slice(0, index-1) + "," + JOBEXTRAUISETTINGS + "," + JOBENFORCEUSERCODE + "}\r\n";
    return result;
}

function setPageProfileDistributionDefault(modelSpec, scriptContext)
{
    try
    {
        if( "RPCS" == GetStringFromProperties(scriptContext.DriverProperties, "pdl") )
        {
            if( modelSpec.isSupported(PRINT_QUALITY_ADJUSTMENT) )
            {
                if( modelSpec.isSupported(COLOR_PROFILE_DISTRIBUTION) &&
                    "default" in modelSpec.modelSpec[COLOR_PROFILE_DISTRIBUTION] )
                {
                    modelSpec.modelSpec[COLOR_PROFILE_DISTRIBUTION]["default"] = "ns0000:Printer";
                }
            }
        }
    }
    catch(e)
    {
    }
}


/**************************************************************
*                                                             *
*              Utility functions                              *
*                                                             *
**************************************************************/

function setPropertyValue(propertyNode, value) {
    //// <summary>
    ////     Set the value contained in the 'Value' node under a 'Property'
    ////     or a 'ScoredProperty' node in the print ticket/print capabilities document.
    //// </summary>
    //// <param name="propertyNode" type="IXMLDOMNode">
    ////     The 'Property'/'ScoredProperty' node.
    //// </param>
    //// <param name="value" type="variant">
    ////     The value to be stored under the 'Value' node.
    //// </param>
    //// <returns type="IXMLDOMNode" mayBeNull="true" locid="R:propertyValue">
    ////     First child 'Property' node if found, Null otherwise.
    //// </returns>
    var valueNode = getPropertyFirstValueNode(propertyNode);
    if (valueNode) {
        var child = valueNode.firstChild;
        if (child) {
            child.nodeValue = value;
            return child;
        }
    }
    return null;
}


function setSubPropertyValue(parentProperty, keywordNamespace, subPropertyName, value) {
    //// <summary>
    ////     Set the value contained in an inner Property node's 'Value' node (i.e. 'Value' node in a Property node
    ////     contained inside another Property node).
    //// </summary>
    //// <param name="parentProperty" type="IXMLDOMNode">
    ////     The parent property node.
    //// </param>
    //// <param name="keywordNamespace" type="String">
    ////     The namespace in which the property name is defined.
    //// </param>
    //// <param name="subPropertyName" type="String">
    ////     The name of the sub-property node.
    //// </param>
    //// <param name="value" type="variant">
    ////     The value to be set in the sub-property node's 'Value' node.
    //// </param>
    //// <returns type="IXMLDOMNode" mayBeNull="true">
    ////     Refer setPropertyValue.
    //// </returns>
    if (!parentProperty ||
        !keywordNamespace ||
        !subPropertyName) {
        return null;
    }
    var subPropertyNode = getProperty(
                            parentProperty,
                            keywordNamespace,
                            subPropertyName);
    return setPropertyValue(
            subPropertyNode,
            value);
}

function getScoredProperty(node, keywordNamespace, scoredPropertyName) {
    //// <summary>
    ////     Retrieve a 'ScoredProperty' element in a print ticket/print capabilities document.
    //// </summary>
    //// <param name="node" type="IXMLDOMNode">
    ////     The scope of the search i.e. the parent node.
    //// </param>
    //// <param name="keywordNamespace" type="String">
    ////     The namespace in which the element's 'name' attribute is defined.
    //// </param>
    //// <param name="scoredPropertyName" type="String">
    ////     The ScoredProperty's 'name' attribute (without the namespace prefix).
    //// </param>
    //// <returns type="IXMLDOMNode" mayBeNull="true">
    ////     The node on success, 'null' on failure.
    //// </returns>

    // Note: It is possible to hard-code the 'psfPrefix' variable in the tag name since the
    // SelectionNamespace property has been set against 'psfPrefix'
    // in validatePrintTicket/completePrintCapabilities.
    return searchByAttributeName(
                node,
                psfPrefix + ":ScoredProperty",
                keywordNamespace,
                scoredPropertyName);
}

function getProperty(node, keywordNamespace, propertyName) {
    //// <summary>
    ////     Retrieve a 'Property' element in a print ticket/print capabilities document.
    //// </summary>
    //// <param name="node" type="IXMLDOMNode">
    ////     The scope of the search i.e. the parent node.
    //// </param>
    //// <param name="keywordNamespace" type="String">
    ////     The namespace in which the element's 'name' attribute is defined.
    //// </param>
    //// <param name="propertyName" type="String">
    ////     The Property's 'name' attribute (without the namespace prefix).
    //// </param>
    //// <returns type="IXMLDOMNode" mayBeNull="true">
    ////     The node on success, 'null' on failure.
    //// </returns>
    return searchByAttributeName(
            node,
            psfPrefix + ":Property",
            keywordNamespace,
            propertyName);
}

function setSelectedOptionName(printSchemaFeature, keywordPrefix, optionName) {
    //// <summary>
    ////      Set the 'name' attribute of a Feature's selected option
    ////      Note: This function should be invoked with Feature type that is retrieved
    ////            via either PrintCapabilties->GetFeature() or PrintTicket->GetFeature().
    ////
    ////      Caution: Setting only the 'name' attribute can result in an invalid option element.
    ////            Some options require their entire subtree to be updated.
    //// </summary>
    //// <param name="printSchemaFeature" type="IPrintSchemaFeature">
    ////     Feature variable.
    //// </param>
    //// <param name="keywordPrefix" type="String">
    ////     The prefix for the optionName parameter.
    //// </param>
    //// <param name="optionName" type="String">
    ////     The name (without prefix) to set as the 'name' attribute.
    //// </param>
    if (!printSchemaFeature ||
        !printSchemaFeature.SelectedOption ||
        !printSchemaFeature.SelectedOption.XmlNode) {
        return;
    }
    printSchemaFeature.SelectedOption.XmlNode.setAttribute(
        "name",
        keywordPrefix + ":" + optionName);
}


/**************************************************************
*                                                             *
*              Functions used by utility functions            *
*                                                             *
**************************************************************/

function getPropertyFirstValueNode(propertyNode) {
    //// <summary>
    ////     Retrieve the first 'value' node found under a 'Property' or 'ScoredProperty' node.
    //// </summary>
    //// <param name="propertyNode" type="IXMLDOMNode">
    ////     The 'Property'/'ScoredProperty' node.
    //// </param>
    //// <returns type="IXMLDOMNode" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!propertyNode) {
        return null;
    }

    var nodeName = propertyNode.nodeName;
    if ((nodeName.indexOf(":Property") < 0) &&
        (nodeName.indexOf(":ScoredProperty") < 0)) {
        return null;
    }

    var valueNode = propertyNode.selectSingleNode(psfPrefix + ":Value");
    return valueNode;
}

function searchByAttributeName(node, tagName, keywordNamespace, nameAttribute) {
    //// <summary>
    ////      Search for a node that with a specific tag name and containing a
    ////      specific 'name' attribute
    ////      e.g. &lt;Bar name=\"ns:Foo\"&gt; is a valid result for the following search:
    ////           Retrieve elements with tagName='Bar' whose nameAttribute='Foo' in
    ////           the namespace corresponding to prefix 'ns'.
    //// </summary>
    //// <param name="node" type="IXMLDOMNode">
    ////     Scope of the search i.e. the parent node.
    //// </param>
    //// <param name="tagName" type="String">
    ////     Restrict the searches to elements with this tag name.
    //// </param>
    //// <param name="keywordNamespace" type="String">
    ////     The namespace in which the element's name is defined.
    //// </param>
    //// <param name="nameAttribute" type="String">
    ////     The 'name' attribute to search for.
    //// </param>
    //// <returns type="IXMLDOMNode" mayBeNull="true">
    ////     IXMLDOMNode on success, 'null' on failure.
    //// </returns>
    if (!node ||
        !tagName ||
        !keywordNamespace ||
        !nameAttribute) {
        return null;
    }

    // For more information on this XPath query, visit:
    // http://blogs.msdn.com/b/benkuhn/archive/2006/05/04/printticket-names-and-xpath.aspx
    var xPathQuery = "descendant::" +
                     tagName +
                     "[substring-after(@name,':')='" +
                     nameAttribute +
                     "']" +
                     "[name(namespace::*[.='" +
                     keywordNamespace +
                     "'])=substring-before(@name,':')]"
    ;

    var nodeRet = node.selectSingleNode(xPathQuery);

    if (nodeRet !== null) {
        return nodeRet;
    }

    return null;
}

function setSelectionNamespace(xmlNode, prefix, namespace) {
    //// <summary>
    ////     This function sets the 'SelectionNamespaces' property on the XML Node.
    ////     For more details: http://msdn.microsoft.com/en-us/library/ms756048(VS.85).aspx
    //// </summary>
    //// <param name="xmlNode" type="IXMLDOMNode">
    ////     The node on which the property is set.
    //// </param>
    //// <param name="prefix" type="String">
    ////     The prefix to be associated with the namespace.
    //// </param>
    //// <param name="namespace" type="String">
    ////     The namespace to be added to SelectionNamespaces.
    //// </param>
    xmlNode.setProperty(
        "SelectionNamespaces",
        "xmlns:" +
            prefix +
            "='" +
            namespace +
            "'"
        );
}


function getPrefixForNamespace(node, namespace) {
    //// <summary>
    ////     This function returns the prefix for a given namespace.
    ////     Example: In 'psf:printTicket', 'psf' is the prefix for the namespace.
    ////     xmlns:psf="http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework"
    //// </summary>
    //// <param name="node" type="IXMLDOMNode">
    ////     A node in the XML document.
    //// </param>
    //// <param name="namespace" type="String">
    ////     The namespace for which prefix is returned.
    //// </param>
    //// <returns type="String">
    ////     Returns the namespace corresponding to the prefix.
    //// </returns>

    if (!node) {
        return null;
    }

    // navigate to the root element of the document.
    var rootNode = node.documentElement;

    // Query to retrieve the list of attribute nodes for the current node
    // that matches the namespace in the 'namespace' variable.
    var xPathQuery = "namespace::node()[.='" +
                namespace +
                "']";
    var namespaceNode = rootNode.selectSingleNode(xPathQuery);
    var prefix = namespaceNode.baseName;

    return prefix;
}


function CreateFeatureSelection(xmlDOMDocInput, strFeatureName, strDisplayName) {
    return CreateFeatureSelection(xmlDOMDocInput, strFeatureName, strDisplayName, pskPrefix);
}

function CreateFeatureSelection(xmlDOMDocInput, strFeatureName, strDisplayName, strPrefix) {
    //// <summary>
    ////     Create a node with name of "strFeatureName", with display name of "strDisplayName"
    //// </summary>
    //// <param name="strFeatureName" type="String">
    ////     The name of the feature node.
    //// <param name="strDisplayName" type="String">
    ////     The display name of the feature node.
    //// </param>
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strFeatureName || strFeatureName.length === 0) {
        return null;
    }

    var featureElement = null;
    featureElement = CreateFeature(xmlDOMDocInput, strFeatureName, strDisplayName, strPrefix);
    if (!featureElement) {
        return null;
    }

    var strPickOne = pskPrefix + ":" + PICKONE_VALUE_NAME;
    var propertyTypeElement = null;
    var strStringType = SCHEMA_XS + ":" + SCHEMA_QNAME;

    propertyTypeElement = CreateFWPropertyWithInput(xmlDOMDocInput, SELECTIONTYPE_VALUE_NAME, strStringType, strPickOne);
    if (propertyTypeElement) {
        featureElement.appendChild(propertyTypeElement);
    }

    return featureElement;
}


function CreateFeature(xmlDOMDocInput, strFeatureNameInput, strDisplayNameInput) {
    return CreateFeature(xmlDOMDocInput, strFeatureNameInput, strDisplayNameInput, pskPrefix);
}


function CreateFeature(xmlDOMDocInput, strFeatureNameInput, strDisplayNameInput, strNSprefix) {
    //// <summary>
    ////     Create a node with name of "strFeatureName", with display name of "strDisplayName"
    //// </summary>
    //// <param name="strFeatureNameInput" type="String">
    ////     The name of the feature node.
    //// <param name="strDisplayNameInput" type="String">
    ////     The display name of the feature node.
    //// </param>
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strFeatureNameInput || strFeatureNameInput.length === 0) {
        return null;
    }

    var strFeature = psfPrefix + ":" + FEATURE_ELEMENT_NAME;
    var strFeatureAttrib = strNSprefix + ":" + strFeatureNameInput;

    var featureElement = null;
    featureElement = CreateXMLElement(xmlDOMDocInput, strFeature, FRAMEWORK_URI);
    if (featureElement) {
        CreateXMLAttribute(xmlDOMDocInput, featureElement, NAME_ATTRIBUTE_NAME, "", strFeatureAttrib);
    }

    if (strDisplayNameInput) {
        var displayProp = CreateFWPropertyWithInput(xmlDOMDocInput, DISPLAYNAME_VALUE_NAME, SCHEMA_STRING, strDisplayNameInput);
        if (displayProp) {
            featureElement.appendChild(displayProp);
        }
    }
    return featureElement;
}


function CreateFWPropertyWithInput(xmlDOMDocInput, strPropNameInput, strTypeInput, strValueInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <param name="strTypeInput" type="String">
    ////     The type of the feature node.
    //// </param>
    //// <param name="strValueInput" type="String">
    ////     The value of the feature node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    var propElement = null;

    propElement = CreateFWProperty(xmlDOMDocInput, strPropNameInput);

    if (propElement) {

        var valueElement = null;

        var strValueElement = psfPrefix + ":" + VALUE_ELEMENT_NAME;
        var strTypeAttribName = SCHEMA_XSI + ":" + SCHEMA_TYPE;
        var strTypeAttribValue = strTypeInput;

        valueElement = CreateXMLElement(xmlDOMDocInput, strValueElement, FRAMEWORK_URI);
        if (valueElement) {
            valueElement.text = strValueInput;
            CreateXMLAttribute(xmlDOMDocInput, valueElement, strTypeAttribName, SCHEMA_INST_URI, strTypeAttribValue);
            propElement.appendChild(valueElement);
        }
    }

    return propElement;
}

function CreateFWProperty(xmlDOMDocInput, strPropNameInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    var propElement = null;
    var strTagName = psfPrefix + ":" + PROPERTY_ELEMENT_NAME;
    var strAttribName = psfPrefix + ":" + strPropNameInput;

    propElement = CreateXMLElement(xmlDOMDocInput, strTagName, FRAMEWORK_URI);
    if (propElement) {
        CreateXMLAttribute(xmlDOMDocInput, propElement, NAME_ATTRIBUTE_NAME, "", strAttribName);
    }

    return propElement;
}

function CreateScoredPropertyRefElement(xmlDOMDocInput, strPropNameInput, strParamRefNameAttribInput) {
    return CreateScoredPropertyRefElement(xmlDOMDocInput, strPropNameInput, strParamRefNameAttribInput, pskPrefix);
}

function CreateScoredPropertyRefElement(xmlDOMDocInput, strPropNameInput, strParamRefNameAttribInput, strPrefix_input) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    if (!strParamRefNameAttribInput || strParamRefNameAttribInput.length === 0) {
        return null;
    }

    var propElement = null;
    var strTagName = psfPrefix + ":" + SCORED_PROP_ELEMENT_NAME;

    var strTagNameUrl = psfPrefix + ":" + psfNs;

    var strAttribName = strPrefix_input + ":" + strPropNameInput;

    propElement = CreateXMLElement(xmlDOMDocInput, strTagName, FRAMEWORK_URI);
    if (propElement) {
        CreateXMLAttribute(xmlDOMDocInput, propElement, NAME_ATTRIBUTE_NAME, "", strAttribName);
    }

    var paramRefElement = null;
    var strParamRefName = psfPrefix + ":" + PARAM_REF_ELEMENT_NAME;
    var strparamAttribName = strPrefix_input + ":" + strParamRefNameAttribInput;

    paramRefElement = CreateXMLElement(xmlDOMDocInput, strParamRefName, FRAMEWORK_URI);
    if (paramRefElement) {
        CreateXMLAttribute(xmlDOMDocInput, paramRefElement, NAME_ATTRIBUTE_NAME, "", strparamAttribName);
    }

    if (propElement && paramRefElement) {
        propElement.appendChild(paramRefElement);
    }

    return propElement;
}


function CreateScoredPropertyIntegerElement(xmlDOMDocInput, strPropNameInput, strParamValue) {
    return CreateScoredPropertyRefElement(xmlDOMDocInput, strPropNameInput, strParamValue, pskPrefix);
}

function CreateScoredPropertyIntegerElement(xmlDOMDocInput, strPropNameInput, strParamValue, strPrefix_input) {
    //// <summary>
    ////     creates a property element of the given integer "strParamValue".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    var propElement = null;
    var strTagName = psfPrefix + ":" + SCORED_PROP_ELEMENT_NAME;

    var strTagNameUrl = psfPrefix + ":" + psfNs;

    var strAttribName = strPrefix_input + ":" + strPropNameInput;

    propElement = CreateXMLElement(xmlDOMDocInput, strTagName, FRAMEWORK_URI);
    if (propElement) {
        CreateXMLAttribute(xmlDOMDocInput, propElement, NAME_ATTRIBUTE_NAME, "", strAttribName);
    }

    var strValueElement = psfPrefix + ":" + VALUE_ELEMENT_NAME;
    var strTypeAttribName = SCHEMA_XSI + ":" + SCHEMA_TYPE;
    var strIntegerType = SCHEMA_XS + ":" + SCHEMA_INTEGER;

    var valueElement = CreateXMLElement(xmlDOMDocInput, strValueElement, FRAMEWORK_URI);
    if (valueElement) {
        valueElement.text = strParamValue;
        CreateXMLAttribute(xmlDOMDocInput, valueElement, strTypeAttribName, SCHEMA_INST_URI, strIntegerType);
    }

    if (propElement && valueElement) {
        propElement.appendChild(valueElement);
    }

    return propElement;
}


function CreateScoredPropertyStringElement(xmlDOMDocInput, strPropNameInput, strParamValue) {
    return CreateScoredPropertyStringElement(xmlDOMDocInput, strPropNameInput, strParamValue, pskPrefix);
}

function CreateScoredPropertyStringElement(xmlDOMDocInput, strPropNameInput, strParamValue, strPrefix_input) {
    //// <summary>
    ////     creates a property element of the given string "strParamValue".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    var propElement = null;
    var strTagName = psfPrefix + ":" + SCORED_PROP_ELEMENT_NAME;

    var strTagNameUrl = psfPrefix + ":" + psfNs;

    var strAttribName = strPrefix_input + ":" + strPropNameInput;

    propElement = CreateXMLElement(xmlDOMDocInput, strTagName, FRAMEWORK_URI);
    if (propElement) {
        CreateXMLAttribute(xmlDOMDocInput, propElement, NAME_ATTRIBUTE_NAME, "", strAttribName);
    }

    var strValueElement = psfPrefix + ":" + VALUE_ELEMENT_NAME;
    var strTypeAttribName = SCHEMA_XSI + ":" + SCHEMA_TYPE;
    var strIntegerType = SCHEMA_XS + ":" + SCHEMA_STRING;

    var valueElement = CreateXMLElement(xmlDOMDocInput, strValueElement, FRAMEWORK_URI);
    if (valueElement) {
        valueElement.text = strParamValue;
        CreateXMLAttribute(xmlDOMDocInput, valueElement, strTypeAttribName, SCHEMA_INST_URI, strIntegerType);
    }

    if (propElement && valueElement) {
        propElement.appendChild(valueElement);
    }

    return propElement;
}



function CreateStringParamDef(xmlDOMDocInput, strPropDefNameInput, bIsPublicKeywordInput, strDefaultValueInput, intMinLengthInput, intMaxLengthInput, strUnitTypeInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropDefNameInput || strPropDefNameInput.length === 0) {
        return null;
    }   

    if (!xmlDOMDocInput) {
        return null;
    }

    var strparamDefName = psfPrefix + ":" + PARAM_DEF_ELEMENT_NAME;
    var strParamAttrib = null;

    if (bIsPublicKeywordInput === false) {
        strParamAttrib = userpskPrefix + ":";
    }
    else {
        strParamAttrib = pskPrefix + ":";
    }

    strParamAttrib += strPropDefNameInput;

    var propDefElement = null;

    propDefElement = CreateXMLElement(xmlDOMDocInput, strparamDefName, FRAMEWORK_URI);
    if (propDefElement) {
        CreateXMLAttribute(xmlDOMDocInput, propDefElement, NAME_ATTRIBUTE_NAME, "", strParamAttrib);
    }

    if (propDefElement) {

        var nodeDataTypeProp = null;
        var nodeDefValProp = null;
        var nodeMaxLengthProp = null;
        var nodeMinLengthProp = null;
        var nodeMandatoryProp = null;
        var nodeUnitTypeProp = null;

        var strQNameType = SCHEMA_XS + ":" + SCHEMA_QNAME;
        var strStringType = SCHEMA_XS + ":" + SCHEMA_STRING;
        var strDataType = DATATYPE_VALUE_NAME;
        var strDefValue = DEFAULTVAL_VALUE_NAME;
        var strMaxLength = MAX_LENGTH_NAME;
        var strIntegerType = SCHEMA_XS + ":" + SCHEMA_INTEGER;
        var strMinLength = MIN_LENGTH_NAME;
        var strMandatory = MANDATORY_VALUE_NAME;
        var strMandatoryValue = pskPrefix + ":" + SCHEMA_CONDITIONAL;
        var strUnitType = UNITTYPE_VALUE_NAME;

        nodeDataTypeProp = CreateFWPropertyWithInput(xmlDOMDocInput, strDataType, strQNameType, strStringType);
        if (nodeDataTypeProp) {
            propDefElement.appendChild(nodeDataTypeProp);
        }

        nodeDefValProp = CreateFWPropertyWithInput(xmlDOMDocInput, strDefValue, strStringType, strDefaultValueInput);
        if (nodeDefValProp) {
            propDefElement.appendChild(nodeDefValProp);
        }

        nodeMaxLengthProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMaxLength, strIntegerType, intMaxLengthInput.toString());
        if (nodeMaxLengthProp) {
            propDefElement.appendChild(nodeMaxLengthProp);
        }

        nodeMinLengthProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMinLength, strIntegerType, intMinLengthInput.toString());
        if (nodeMinLengthProp) {
            propDefElement.appendChild(nodeMinLengthProp);
        }

        nodeMandatoryProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMandatory, strQNameType, strMandatoryValue);
        if (nodeMandatoryProp) {
            propDefElement.appendChild(nodeMandatoryProp);
        }

        nodeUnitTypeProp = CreateFWPropertyWithInput(xmlDOMDocInput, strUnitType, strStringType, strUnitTypeInput);
        if (nodeUnitTypeProp) {
            propDefElement.appendChild(nodeUnitTypeProp);
        }

    }

    return propDefElement;
}

function CreateIntParamDef(xmlDOMDocInput, strPropDefNameInput, bIsPublicKeywordInput, intDefaultValueInput, intMinValueInput, intMaxValueInput, intMultipleInout, strUnitTypeInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropDefNameInput || strPropDefNameInput.length === 0) {
        return null;
    }

    if (!xmlDOMDocInput) {
        return null;
    }

    var strparamDefName = psfPrefix + ":" + PARAM_DEF_ELEMENT_NAME;
    var strParamAttrib = null;

    if (bIsPublicKeywordInput === false) {
        strParamAttrib = userpskPrefix + ":";
    }
    else {
        strParamAttrib = pskPrefix + ":";
    }

    strParamAttrib += strPropDefNameInput;
    var propDefElement = null;

    propDefElement = CreateXMLElement(xmlDOMDocInput, strparamDefName, FRAMEWORK_URI);
    if (propDefElement) {
        CreateXMLAttribute(xmlDOMDocInput, propDefElement, NAME_ATTRIBUTE_NAME, "", strParamAttrib);
    }

    if (propDefElement) {

        var nodeDataTypeProp = null;
        var nodeDefValProp = null;
        var nodeMaxValueProp = null;
        var nodeMinValueProp = null;
        var nodeMultipleProp = null;
        var nodeMandatoryProp = null;
        var nodeUnitTypeProp = null;

        var strQNameType = SCHEMA_XS + ":" + SCHEMA_QNAME;
        var strStringType = SCHEMA_XS + ":" + SCHEMA_STRING;
        var strDataType = DATATYPE_VALUE_NAME;
        var strDefValue = DEFAULTVAL_VALUE_NAME;
        var strMaxValue = MAX_VALUE_NAME;
        var strIntegerType = SCHEMA_XS + ":" + SCHEMA_INTEGER;
        var strMinValue = MIN_VALUE_NAME;
        var strMultiple = MULTIPLE_VALUE_NAME;
        var strMandatory = MANDATORY_VALUE_NAME;
        var strMandatoryValue = pskPrefix + ":" + SCHEMA_CONDITIONAL;
        var strUnitType = UNITTYPE_VALUE_NAME;        

        nodeDataTypeProp = CreateFWPropertyWithInput(xmlDOMDocInput, strDataType, strQNameType, strIntegerType);
        if (nodeDataTypeProp) {
            propDefElement.appendChild(nodeDataTypeProp);
        }

        nodeDefValProp = CreateFWPropertyWithInput(xmlDOMDocInput, strDefValue, strIntegerType, intDefaultValueInput.toString());
        if (nodeDefValProp) {
            propDefElement.appendChild(nodeDefValProp);
        }

        nodeMaxValueProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMaxValue, strIntegerType, intMaxValueInput.toString());
        if (nodeMaxValueProp) {
            propDefElement.appendChild(nodeMaxValueProp);
        }

        nodeMinValueProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMinValue, strIntegerType, intMinValueInput.toString());
        if (nodeMinValueProp) {
            propDefElement.appendChild(nodeMinValueProp);
        }

        nodeMultipleProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMultiple, strIntegerType, intMultipleInout.toString());
        if (nodeMultipleProp) {
            propDefElement.appendChild(nodeMultipleProp);
        }

        nodeMandatoryProp = CreateFWPropertyWithInput(xmlDOMDocInput, strMandatory, strQNameType, strMandatoryValue);
        if (nodeMandatoryProp) {
            propDefElement.appendChild(nodeMandatoryProp);
        }

        nodeUnitTypeProp = CreateFWPropertyWithInput(xmlDOMDocInput, strUnitType, strStringType, strUnitTypeInput);
        if (nodeUnitTypeProp) {
            propDefElement.appendChild(nodeUnitTypeProp);
        }

    }

    return propDefElement;
}

function CreateXMLElement(xmlDOMDocInput, strElementNameInput, strTargetURIInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strElementNameInput" type="String">
    ////     The name of the element node.
    //// <param name="strTargetURIInput" type="String">
    ////     The name of the Target URI node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>

    //if (!strElementNameInput || strElementNameInput.length === 0) {
    if (null === strElementNameInput) {
        return null;
    }

    var currentNode = null;
    currentNode = xmlDOMDocInput.XmlNode.createNode(1, strElementNameInput, strTargetURIInput);

    if (currentNode) {
        return currentNode;
    }

    return null;

}

function CreateXMLAttribute(xmlDOMDocInput, elementInput, strNameInput, strTargetURIInput, strValueInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="elementInput" type="IXMLDOMElement">
    ////     The element node.
    //// <param name="strNameInput" type="String">
    ////     The name of the Attribute node.
    //// <param name="strTargetURIInput" type="String">
    ////     The name of the Target URI node.
    //// <param name="strValueInput" type="String">
    ////     The value of the Attribute node.

    if (!strNameInput || strNameInput.length === 0) {
        return null;
    }

    var currentNode = null;
    var currentAttributeNode = null;

    var nodeType = 2;
    currentNode = xmlDOMDocInput.XmlNode.createNode(nodeType, strNameInput, strTargetURIInput);
    
    if (currentNode) {
        if (strValueInput) {
            currentNode.value = strValueInput;
            elementInput.setAttributeNode(currentNode);
        }
    }

}



function CreateOption(xmlDOMDocInput, strOptionNameInput, strDisplayNameInput) {

    return CreateOption(xmlDOMDocInput, strOptionNameInput, strDisplayNameInput, pskPrefix);
}

function CreateOption(xmlDOMDocInput, strOptionNameInput, strDisplayNameInput, strNameSpace) {
    //// <summary>
    ////     Create a node with name of "strFeatureName", with display name of "strDisplayName"
    //// </summary>
    //// <param name="strOptionNameInput" type="String">
    ////     The name of the option node.
    //// <param name="strDisplayNameInput" type="String">
    ////     The display name of the feature node.
    //// </param>
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>

    if (null === strOptionNameInput) {
        return null;
    }

    var strOption = psfPrefix + ":" + OPTION_ELEMENT_NAME;
    var strOptionAttrib = (strOptionNameInput === "") ? strOptionNameInput : strNameSpace + ":" + strOptionNameInput;

    var pOptionElement = null;
    pOptionElement = CreateXMLElement(xmlDOMDocInput, strOption, FRAMEWORK_URI);
    if (pOptionElement) {
        CreateXMLAttribute(xmlDOMDocInput, pOptionElement, NAME_ATTRIBUTE_NAME, FRAMEWORK_URI, strOptionAttrib);
    }

    if (strDisplayNameInput) {
        var pDisplayPropElem = CreateFWPropertyWithInput(xmlDOMDocInput, DISPLAYNAME_VALUE_NAME, SCHEMA_STRING, strDisplayNameInput);
        if (pOptionElement && pDisplayPropElem) {
            pOptionElement.appendChild(pDisplayPropElem);
        }
    }

    return pOptionElement;

}

function CreateIntParamRefElement(devModeProperties, printTicket, scriptContext, strParamRefName) {

    var m_pRootDocument = printTicket;
    var strTypeName = null;
    var strValue = null;

    strTypeName = SCHEMA_XS + ":" + SCHEMA_STRING;
    try {
        strValue = devModeProperties.GetString(strParamRefName);
        var nodeParamIniValue = getParamDefInitNode(m_pRootDocument.XmlNode.documentElement, false, strParamRefName);
        if (!nodeParamIniValue ) {
            nodeParamIniValue = CreateIntParamRefIni(m_pRootDocument, strParamRefName, false, strTypeName, strValue.toString());
            if (nodeParamIniValue) {
                printTicket.XmlNode.lastChild.appendChild(nodeParamIniValue);
            }
        }
        else {
            nodeParamIniValue.firstChild.text = strValue.toString();
        }
    }
    catch (e) {
        var defaultValueKey = "Default" + strParamRefName + "Value";
        strValue = scriptContext.DriverProperties.GetString(defaultValueKey);
        var nodeParamIniValue = CreateIntParamRefIni(m_pRootDocument, strParamRefName, false, strTypeName, strValue.toString);
        if (nodeParamIniValue) {
            printTicket.XmlNode.lastChild.appendChild(nodeParamIniValue);
        }
    }
 
}


function getParamDefInitNode(xmlSrcNode_Input, boolIsPublicKeyword, strOptionName_Input) {

    if (!xmlSrcNode_Input) {
        return null;
    }

    if (!strOptionName_Input || strOptionName_Input.length === 0) {
        return null;
    }

    var strNameSpace = null;

    if (boolIsPublicKeyword === false) {
        strNameSpace = ricohNs;
    }
    else {
        strNameSpace = pskNs;
    }

    var xmlNode_Out = searchByAttributeName(xmlSrcNode_Input, psfPrefix + ":" + PARAM_INIT_ELEMENT_NAME, strNameSpace, strOptionName_Input);

    return xmlNode_Out;
}


function CreateIntParamRefIni(xmlDOMDocInput, strPropDefNameInput, bIsPublicKeywordInput, strTypeInput, strValueInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropDefNameInput || strPropDefNameInput.length === 0) {
        return null;
    }

    if (!xmlDOMDocInput) {
        return null;
    }

    var strparamDefName = psfPrefix + ":" + PARAM_INIT_ELEMENT_NAME;
    var strParamAttrib = null;

    if (bIsPublicKeywordInput === false) {
        strParamAttrib = userpskPrefix + ":";
    } else {
        strParamAttrib = pskPrefix + ":";
    }

    strParamAttrib += strPropDefNameInput;

    var propDefElement = null;

    propDefElement = CreateXMLElement(xmlDOMDocInput, strparamDefName, FRAMEWORK_URI);
    if (propDefElement) {
        CreateXMLAttribute(xmlDOMDocInput, propDefElement, NAME_ATTRIBUTE_NAME, FRAMEWORK_URI, strParamAttrib);
    }

    if (propDefElement) {

        var valueElement = null;

        var strValueElement = psfPrefix + ":" + VALUE_ELEMENT_NAME;
        var strTypeAttribName = SCHEMA_XSI + ":" + SCHEMA_TYPE;
        var strAttribValue = strTypeInput;

        valueElement = CreateXMLElement(xmlDOMDocInput, strValueElement, FRAMEWORK_URI);
        if (valueElement) {
            valueElement.text = strValueInput;
            CreateXMLAttribute(xmlDOMDocInput, valueElement, strTypeAttribName, SCHEMA_INST_URI, strAttribValue);
            propDefElement.appendChild(valueElement);
        }
    }

    return propDefElement;
}


function GetUnicodeToObject(drvProperties, loadFileName) {
    var bytesData = GetBytesFromProperties(drvProperties, loadFileName);
    var multiChar = WideCharBytesToMultiCharByte(bytesData);
    var object = ConvertToJsonObjectFromBytes(multiChar);

    return object;
}

function CreatelangPropertyWithInput(xmlDOMDocInput, strPropNameInput, strTypeInput, strValueInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <param name="strTypeInput" type="String">
    ////     The type of the feature node.
    //// </param>
    //// <param name="strValueInput" type="String">
    ////     The value of the feature node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    var propElement = null;

    propElement = CreatelangProperty(xmlDOMDocInput, strPropNameInput);

    if (propElement) {

        var valueElement = null;

        var strValueElement = psfPrefix + ":" + VALUE_ELEMENT_NAME;
        var strTypeAttribName = SCHEMA_XSI + ":" + SCHEMA_TYPE;
        var strTypeAttribValue = SCHEMA_XS + ":" + strTypeInput;

        valueElement = CreateXMLElement(xmlDOMDocInput, strValueElement, FRAMEWORK_URI);
        if (valueElement) {
            valueElement.text = strValueInput;
            CreateXMLAttribute(xmlDOMDocInput, valueElement, strTypeAttribName, SCHEMA_INST_URI, strTypeAttribValue);
            propElement.appendChild(valueElement);
        }
    }

    return propElement;
}

function CreatelangProperty(xmlDOMDocInput, strPropNameInput) {
    //// <summary>
    ////     creates a property element of the given name "strPropNameInput".
    //// </summary>
    //// <param name="strPropNameInput" type="String">
    ////     The name of the property node.
    //// <returns type="IXMLDOMElement" mayBeNull="true">
    ////     The 'Value' node on success, 'null' on failure.
    //// </returns>
    if (!strPropNameInput || strPropNameInput.length === 0) {
        return null;
    }

    var propElement = null;
    var strTagName = psfPrefix + ":" + PROPERTY_ELEMENT_NAME;
    var strAttribName = pskPrefix + ":" + strPropNameInput;

    propElement = CreateXMLElement(xmlDOMDocInput, strTagName, FRAMEWORK_URI);
    if (propElement) {
        CreateXMLAttribute(xmlDOMDocInput, propElement, NAME_ATTRIBUTE_NAME, FRAMEWORK_URI, strAttribName);
    }

    return propElement;
}


// String(UTF-8) to Base64 String
function EncodeBase64FromUTF8(str)
{
    var nums = ConvertNumArray(str);
    return EncodeBase64(nums);
}

// String(UTF-16) to Base64 String
function EncodeBase64FromUTF16(str) {
    str = unescape(encodeURIComponent(str)); // convert from UTF-16 to UTF-8
    return EncodeBase64FromUTF8(str);
}

// String to Numeric Array
function ConvertNumArray(str) {
    var numerics = [];
    for (var i = 0; i < str.length; i++) {
        numerics.push(str.charAt(i).charCodeAt(0));
    }
    return numerics;
}

// Numeric Array to Base64 String
function EncodeBase64(numArray)
{
    var base64chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    var encodeTable = base64chars.split(""); // converted to an array of one character

    // Convert every 3 bytes
    var padding = 3 - (numArray.length % 3);
    if (3 > padding)
    {
        for (var j = 0; j < padding; j++)
        {
            numArray[numArray.length] = 0; // Add 0 padding
        }
    }
    else
    {
        padding = 0;
    }

    var base64 = "";
    for (var i = 0; i < numArray.length; i += 3)
    {
        // 3 bytes concatenated (24bit)
        var block = (numArray[i] << 16) | (numArray[i + 1] << 8) | numArray[i + 2];
        // Convert to ASCII code every 6 bits
        base64 += encodeTable[(block >> 18) & 0x3f]
            + encodeTable[(block >> 12) & 0x3f]
            + encodeTable[(block >> 6) & 0x3f]
            + encodeTable[block & 0x3f];
    }

    // Padding "="
    if (0 < padding)
    {
        base64 = base64.substr(0, base64.length - padding);
        for (var h = 0; h < padding; h++) {
            base64 += "=";
        }
    }

    return base64;
}

