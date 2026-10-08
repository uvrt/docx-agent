"""WordprocessingML's vocabulary, registered with ooxml-edit.

ooxml-edit knows only the packaging namespaces every format shares; this module teaches it
Word's prefixes and the child sequences of the elements docx-agent inserts into, so that
:func:`ooxml_edit.xml.insert_in_order` puts a new child where the schema requires it (Word
answers a misordered ``w:pPr`` with a repair prompt, not an error message).

The sequences are ECMA-376 Part 1 (transitional), section 17, with the MS-DOCX extensions
sorting last, as Word writes them.  Only parents something here inserts into need an entry;
the rest are listed because E1 will, and because the validity checks read them.
"""

from __future__ import annotations

from ooxml_edit.xml import (  # noqa: F401  (re-exported)
    CHILD_ORDER,
    NAMESPACES,
    Element,
    append_in_order,
    child_elements,
    find,
    findall,
    get_int,
    insert_in_order,
    local_name,
    make,
    parse_xml,
    prefixed_name,
    qn,
    register_child_order,
    register_namespaces,
    remove,
    serialize,
    set_attr,
    subelement,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

register_namespaces({
    "w": W,
    "w14": W14,
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
})

#: ``xml:space``, which a ``w:t`` needs when its text starts or ends with a space.
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# -- child sequences -------------------------------------------------------------------------

#: Everything that may sit among a paragraph's runs (EG_PContent with the range markup and
#: revision containers): one rank, since they interleave freely.
_PARAGRAPH_CONTENT = (
    "w:r", "w:hyperlink", "w:fldSimple", "w:customXml", "w:smartTag", "w:sdt", "w:dir",
    "w:bdo", "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:proofErr", "w:permStart",
    "w:permEnd", "w:bookmarkStart", "w:bookmarkEnd", "w:moveFromRangeStart",
    "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd", "w:commentRangeStart",
    "w:commentRangeEnd", "w:customXmlInsRangeStart", "w:customXmlInsRangeEnd",
    "w:customXmlDelRangeStart", "w:customXmlDelRangeEnd", "w:customXmlMoveFromRangeStart",
    "w:customXmlMoveFromRangeEnd", "w:customXmlMoveToRangeStart", "w:customXmlMoveToRangeEnd",
    "m:oMathPara", "m:oMath", "w:subDoc",
)

#: Block-level content (EG_BlockLevelElts with range markup): one rank.
_BLOCK_CONTENT = (
    "w:p", "w:tbl", "w:sdt", "w:customXml", "w:altChunk", "w:proofErr", "w:permStart",
    "w:permEnd", "w:bookmarkStart", "w:bookmarkEnd", "w:moveFromRangeStart",
    "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd", "w:commentRangeStart",
    "w:commentRangeEnd", "w:customXmlInsRangeStart", "w:customXmlInsRangeEnd",
    "w:customXmlDelRangeStart", "w:customXmlDelRangeEnd", "w:customXmlMoveFromRangeStart",
    "w:customXmlMoveFromRangeEnd", "w:customXmlMoveToRangeStart", "w:customXmlMoveToRangeEnd",
    "w:ins", "w:del", "w:moveFrom", "w:moveTo", "m:oMathPara", "m:oMath",
)

#: EG_RangeMarkupElements and the proofing and permission marks around them: one rank.
_RANGE_MARKUP = (
    "w:proofErr", "w:permStart", "w:permEnd", "w:bookmarkStart", "w:bookmarkEnd",
    "w:moveFromRangeStart", "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd",
    "w:commentRangeStart", "w:commentRangeEnd", "w:customXmlInsRangeStart",
    "w:customXmlInsRangeEnd", "w:customXmlDelRangeStart", "w:customXmlDelRangeEnd",
    "w:customXmlMoveFromRangeStart", "w:customXmlMoveFromRangeEnd",
    "w:customXmlMoveToRangeStart", "w:customXmlMoveToRangeEnd", "w:ins", "w:del",
    "w:moveFrom", "w:moveTo",
)

#: EG_RunInnerContent: one rank.
_RUN_CONTENT = (
    "w:br", "w:t", "w:contentPart", "w:delText", "w:instrText", "w:delInstrText",
    "w:noBreakHyphen", "w:softHyphen", "w:dayShort", "w:monthShort", "w:yearShort",
    "w:dayLong", "w:monthLong", "w:yearLong", "w:annotationRef", "w:footnoteRef",
    "w:endnoteRef", "w:separator", "w:continuationSeparator", "w:sym", "w:pgNum", "w:cr",
    "w:tab", "w:object", "w:pict", "w:fldChar", "w:ruby", "w:footnoteReference",
    "w:endnoteReference", "w:commentReference", "w:drawing", "w:ptab",
    "w:lastRenderedPageBreak", "mc:AlternateContent",
)

#: EG_RPrBase, in order.
_RPR_BASE = (
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
    "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof",
    "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern",
    "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
    "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
    "w:specVanish", "w:oMath",
)

#: CT_PPrBase, in order.
_PPR_BASE = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
    "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
    "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
    "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
    "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle",
)

_SETTINGS = (
    "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation",
    "w:removeDateAndTime", "w:doNotDisplayPageBoundaries", "w:displayBackgroundShape",
    "w:printPostScriptOverText", "w:printFractionalCharacterWidth", "w:printFormsData",
    "w:embedTrueTypeFonts", "w:embedSystemFonts", "w:saveSubsetFonts", "w:saveFormsData",
    "w:mirrorMargins", "w:alignBordersAndEdges", "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter", "w:gutterAtTop", "w:hideSpellingErrors",
    "w:hideGrammaticalErrors", "w:activeWritingStyle", "w:proofState", "w:formsDesign",
    "w:attachedTemplate", "w:linkStyles", "w:stylePaneFormatFilter",
    "w:stylePaneSortMethod", "w:documentType", "w:mailMerge", "w:revisionView",
    "w:trackRevisions", "w:doNotTrackMoves", "w:doNotTrackFormatting",
    "w:documentProtection", "w:autoFormatOverride", "w:styleLockTheme", "w:styleLockQFSet",
    "w:defaultTabStop", "w:autoHyphenation", "w:consecutiveHyphenLimit",
    "w:hyphenationZone", "w:doNotHyphenateCaps", "w:showEnvelope", "w:summaryLength",
    "w:clickAndTypeStyle", "w:defaultTableStyle", "w:evenAndOddHeaders",
    "w:bookFoldRevPrinting", "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin", "w:doNotShadeFormData", "w:noPunctuationKerning",
    "w:characterSpacingControl", "w:printTwoOnOne", "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter", "w:noLineBreaksBefore", "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema", "w:saveInvalidXml", "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText", "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly",
    "w:useXSLTWhenSaving", "w:saveThroughXslt", "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace", "w:updateFields", "w:hdrShapeDefaults",
    "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids", "m:mathPr",
    "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "w:schemaLibrary",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
)

#: Revision marks a paragraph mark's run properties carry before its formatting.
_MARK_REVISIONS = ("w:ins", "w:del", "w:moveFrom", "w:moveTo")

register_child_order({
    "w:body": (_BLOCK_CONTENT, "w:sectPr"),
    "w:p": ("w:pPr", _PARAGRAPH_CONTENT),
    "w:r": ("w:rPr", _RUN_CONTENT),
    "w:pPr": _PPR_BASE + ("w:rPr", "w:sectPr", "w:pPrChange"),
    "w:rPr": _MARK_REVISIONS + _RPR_BASE + ("w:rPrChange",),
    "w:sectPr": (
        ("w:headerReference", "w:footerReference"), "w:footnotePr", "w:endnotePr", "w:type",
        "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType",
        "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
        "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
    ),
    "w:tbl": (_RANGE_MARKUP, "w:tblPr", "w:tblGrid", ("w:tr", "w:customXml", "w:sdt") + _RANGE_MARKUP),
    "w:tblPr": (
        "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
        "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd",
        "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
        "w:tblDescription", "w:tblPrChange",
    ),
    "w:tblGrid": ("w:gridCol", "w:tblGridChange"),
    "w:tr": ("w:tblPrEx", "w:trPr", ("w:tc", "w:customXml", "w:sdt") + _RANGE_MARKUP),
    # CT_TrPrBase is a repeating choice, so its members share a rank.
    "w:trPr": (
        ("w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore", "w:wAfter",
         "w:cantSplit", "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden"),
        "w:ins", "w:del", "w:trPrChange",
    ),
    "w:tc": ("w:tcPr", _BLOCK_CONTENT),
    "w:tcPr": (
        "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd",
        "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
        "w:headers", ("w:cellIns", "w:cellDel", "w:cellMerge"), "w:tcPrChange",
    ),
    "w:tblPrEx": (
        "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
        "w:tblCellMar", "w:tblLook", "w:tblPrExChange",
    ),
    # E5: the sides of borders and margins, as CT_TblBorders, CT_TcBorders and CT_TblCellMar
    # order them (``start``/``end`` are the transitional spellings' strict twins).
    "w:tblBorders": ("w:top", ("w:left", "w:start"), "w:bottom", ("w:right", "w:end"), "w:insideH", "w:insideV"),
    "w:tcBorders": ("w:top", ("w:left", "w:start"), "w:bottom", ("w:right", "w:end"), "w:insideH", "w:insideV",
                    "w:tl2br", "w:tr2bl"),
    "w:tcMar": ("w:top", ("w:left", "w:start"), "w:bottom", ("w:right", "w:end")),
    "w:tblCellMar": ("w:top", ("w:left", "w:start"), "w:bottom", ("w:right", "w:end")),
    "w:sdt": ("w:sdtPr", "w:sdtEndPr", "w:sdtContent"),
    # CT_SdtPr, as Word writes it (measured: tag before id); the kinds last, one rank.
    "w:sdtPr": (
        "w:rPr", "w:alias", "w:tag", "w:id", "w:lock", "w:placeholder", "w:temporary", "w:showingPlcHdr",
        "w:dataBinding", "w:label", "w:tabIndex",
        ("w:equation", "w:comboBox", "w:date", "w:docPartObj", "w:docPartList", "w:dropDownList", "w:picture",
         "w:richText", "w:text", "w:citation", "w:group", "w:bibliography", "w14:checkbox", "w15:color",
         "w15:appearance", "w15:dataBinding", "w15:webExtensionLinked", "w15:webExtensionCreated",
         "w15:repeatingSection", "w15:repeatingSectionItem"),
    ),
    "wp:anchor": (
        "wp:simplePos", "wp:positionH", "wp:positionV", "wp:extent", "wp:effectExtent",
        ("wp:wrapNone", "wp:wrapSquare", "wp:wrapTight", "wp:wrapThrough", "wp:wrapTopAndBottom"),
        "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic", "wp14:sizeRelH", "wp14:sizeRelV",
    ),
    "wp:inline": ("wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"),
    "w:hdr": (_BLOCK_CONTENT,),
    "w:ftr": (_BLOCK_CONTENT,),
    "w:footnote": (_BLOCK_CONTENT,),
    "w:endnote": (_BLOCK_CONTENT,),
    "w:comment": (_BLOCK_CONTENT,),
    "w:txbxContent": (_BLOCK_CONTENT,),
    "w:settings": _SETTINGS,
    "w:numPr": ("w:ilvl", "w:numId", "w:numberingChange", "w:ins"),
    # CT_FtnProps / CT_EdnProps, and in the settings CT_FtnDocProps / CT_EdnDocProps, whose
    # separator notes follow.
    "w:footnotePr": ("w:pos", "w:numFmt", "w:numStart", "w:numRestart", "w:footnote"),
    "w:endnotePr": ("w:pos", "w:numFmt", "w:numStart", "w:numRestart", "w:endnote"),
    "w:numbering": ("w:numPicBullet", "w:abstractNum", "w:num", "w:numIdMacAtCleanup"),
    "w:abstractNum": ("w:nsid", "w:multiLevelType", "w:tmpl", "w:name", "w:styleLink",
                      "w:numStyleLink", "w:lvl"),
    "w:num": ("w:abstractNumId", "w:lvlOverride"),
    "w:lvlOverride": ("w:startOverride", "w:lvl"),
    "w:lvl": ("w:start", "w:numFmt", "w:lvlRestart", "w:pStyle", "w:isLgl", "w:suff",
              "w:lvlText", "w:lvlPicBulletId", "w:legacy", "w:lvlJc", "w:pPr", "w:rPr"),
    "w:hyperlink": (_PARAGRAPH_CONTENT,),
    "w:style": (
        "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine", "w:hidden",
        "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked",
        "w:personal", "w:personalCompose", "w:personalReply", "w:rsid", "w:pPr", "w:rPr",
        "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr",
    ),
})

#: Elements whose run properties a paragraph mark's ``w:rPr`` holds as revisions, and the
#: property-change records: neither is formatting a new run should inherit.
REVISION_PROPERTY_TAGS = frozenset(qn(t) for t in _MARK_REVISIONS + ("w:rPrChange",))
