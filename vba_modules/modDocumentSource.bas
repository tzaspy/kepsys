Attribute VB_Name = "modDocumentSource"


' From the DynapPDF HELP file!
'SetMetaConvFlags
'Syntax:
'SI32 pdfSetMetaConvFlags(
'const void* IPDF, // Instance pointer
'TMetaFlags Flags) // see below

'typedef UI32 TMetaFlags;
'#define mfDefault 0x00000 // No flags
'#define mfDebug 0x00001 // Insert debug comments
'#define mfShowBounds 0x00002 // Show the bounding boxes of text
'#define mfNoTextScaling 0x00004 // Do not scale text records
'#define mfClipView 0x00008 // Clip the output rectangle
'#define mfUseRclBounds 0x00010 // Use the raw bounding box rclBounds
'#define mfNoClippingRgn 0x00040 // Ignore clipping regions
'#define mfNoFontEmbedding 0x00080 // Don't embed fonts used by the EMF
'#define mfNoImages 0x00100 // Ignore image records
'#define mfNoStdPatterns 0x00200 // Ignore standard hatch patterns
'#define mfNoBmpPatterns 0x00400 // Ignore bitmap patterns
'#define mfNoText 0x00800 // Ignore text records
'#define mfUseUnicode 0x01000 // Use always Unicode to print text
'#define mfUseTextScaling 0x04000 // Scale text (see description)
'#define mfNoUnicode 0x08000 // Avoid the usage of Unicode fonts
'#define mfFullScale 0x10000 // Scale coordinates to Windows size
'#define mfUseRclFrame 0x20000 // See description
'#define mfDefBkModeTransp 0x40000 // Initial backg. mode is transparent
'#define mfApplyBidiAlgo 0x80000 // Apply the bidirectional algorithm
'#define mfGDIFontSelection 0x100000 // Use the GDI to select fonts
'// Obsolete flags -> These flags are ignored, do not longer use them!
'#define mfUseSpacingArray 0x0020 // Enabled by default

'Function Reference Page 412 of 473
'#define mfIntersectClipRect 0x2000 // Enabled by default
'The function sets specific flags to control the conversion of metafiles. The flags are described in detail
'on the next page. This function cannot fail the return value is always 1.
'flag Description
'mfDefault This is the default behaviour. No specific parameters are used for
'metafile conversion.

'mfDebug If set, the EMF record names are printed to the content stream which
'produces a specific output. Open the PDF file in a good text editor
'such as Textpad to view the output. The PDF file must not be
'compressed if this flag is used, otherwise you can't see the debug
'strings. Compression can be disabled with the function
'SetCompressionLevel().

'mfShowBounds If set, the bounding boxes of text strings are shown by inserting a
'stroked rectangle. Use this flag if text strings appear misplaced.

'mfNoTextScaling If set, text strings are not scaled and no kerning space will be applied.
'EMF files contain sometimes an invalid spacing array especially when
'the original string was substituted from an Arabic code page. In such
'cases, characters do overlap. To avoid this effect disable the usage of
'the intercharacter spacing array; all strings are then printed without
'scaling or kerning space.

'mfClipView If set, the metafile is drawn into a clipping rectangle in the size of the
'metafile or currently defined view. This flag should always be set if a
'user defined view is used. See InsertMetafileExt() for further
'information.

'mfUseRclBounds If set, the value of rclFrame is not used to calculate the size and
'position of the EMF file. rclFrame is a member of the EMF's header
'that specifies the logical output size in 0.01 millimetres. However, this
'rectangle is changed if the file will be rendered onto a monitor DC or
'printer DC. To get consistent output results the output frame can be
'ignored. The graphic must then be placed manually onto the output
'page because the rectangle rclBounds represents the visible extent of
'the picture only (the graphic is drawn without a border).
'Function Reference Page 413 of 473
'flag Description

'mfNoClippingRgn If set, clipping regions will be ignored.

'mfNoFontEmbedding If set, fonts used by an EMF file are not embedded.

'mfNoImages If set, image records are ignored.

'mfNoStdPatterns If set, standard hatch patterns are not applied.

'mfNoBmpPatterns If set, bitmap patterns are ignored.

'mfNoText If set, text records are ignored.

'mfUseUnicode If set, the character set within CreateFont() records will be ignored and
'all strings are printed in Unicode mode (EMF files contain Unicode
'strings only). This flag can be used to avoid the conversion of strings
'to the ANSI character set if ANSI_CHARSET was used in the
'CreateFont() record. The character set is often wrongly defined in EMF
'files so that characters outside of the ANSI_CHARSET are replaced by
'question marks due to the default conversion to ANSI if the character
'set ANSI_CHARSET is used.

'mfUseTextScaling If set, strings are scaled instead of applying kerning space to get the
'correct string width. Text scaling produces often better results due to
'limited precision of the integer values of the intercharacter spacing
'array. However, text scaling cannot always be used because characters
'can be placed individually on the x-axis by applying kerning space. In
'the latter case, strings can overlap or single characters appear on a
'wrong x-coordinate; do not set this flag in this case.

'mfNoUnicode If set, all strings are converted to the code page 1252 and no Unicode
'font will be embedded in the PDF file during EMF conversion. This
'flag should be set if PDF-1.2 compatibility is recommended because
'PDF-1.2 does not support Unicode fonts. However, note that you get
'invalid results if the EMF file contains characters outside the code
'Page 1252#

'mfFullScale If set, all coordinates are scaled to the output window size. Set this
'flag if the EMF file uses 32 bit coordinates, e.g. large CAD drawings.
'Full scaling avoids floating point overflows in PDF viewer
'applications because all coordinate transformations are already
'applied. The resulting file size is also smaller due to the smaller
'coordinate values which must be stored in the PDF file.
'Function Reference Page 414 of 473
'flag Description

'mfUseRclFrame If set, the rectangle rclFrame of the EMF header is used to calculate the
'picture size. This flag is primarily used to convert EMF files which
'were originally created from non-portable WMF files. Set this flag if
'the EMF picture appears wrongly scaled.

'mfDefBkModeTransp If set, the initial background mode is set to transparent. SetBkMode
'records still override this state. The default background mode is
'opaque in the GDI. This state causes that a rectangle is printed in
'background of any text string, also if the rectangle is not required.
'Especially if text strings are printed as single characters the opaque
'background can drastically increase the resulting file size due to the
'many rectangles. This flag should always be set if the EMF or WMF
'files do not initialize the background mode to the required value.

'mfApplyBidiAlgo If set, the bidirectional algorithm is applied on Unicode strings. This
'flag must be set to process Hebrew text correctly.

'mfGDIFontSelection If set, DynaPDF uses a GDI device context to select fonts. This flag can
'be set to make sure that DynaPDF selects exactly the fonts which the
'GDI uses to render to EMF file.


