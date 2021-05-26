using System;
using System.IO;
using System.IO.Compression;

namespace H2O_odf
{
    
    class OdfManager
    { 
        public void makeOdf()
        {
            string startpath = @"C:\Users\park\source\repos\H2O_odf\H2O_odf\test";
            string zippath = @"C:\Users\park\source\repos\H2O_odf\H2O_odf\test.odt";

            DirectoryInfo activeDir = new DirectoryInfo(startpath);
            int i = 0;

            while (activeDir.Exists)
            {
                i++;
                startpath = startpath + "(" + i.ToString() + ")";
                activeDir = new DirectoryInfo(startpath);

            }
            activeDir.Create();

            Content content = new Content(startpath);
            mamifest mamifest = new mamifest(startpath);
            Meta meta = new Meta(startpath);
            Mimetype mimetype = new Mimetype(startpath);
            Settings setting = new Settings(startpath);
            Styles style = new Styles(startpath);

            content.writeOdf();
            mamifest.writeOdf();
            meta.writeOdf();
            mimetype.writeOdf();
            setting.writeOdf();
            style.writeOdf();

            Configuration conf = new Configuration(startpath);
            Metainf meta_inf = new Metainf(startpath);
            Thumbnails thumbnail = new Thumbnails(startpath);

            conf.createDirectory();
            meta_inf.createDirectory();
            thumbnail.createDirectory();

            try
            {
                ZipFile.CreateFromDirectory(startpath, zippath);
            }
            catch(IOException){
                File.Delete(zippath);
                ZipFile.CreateFromDirectory(startpath, zippath);
            }


            //activeDir.Delete(true);


        }

   
    }
    class Content
    {
        private string file;//경로
        private int count = 1;//단락 수
        private contentDetail detail = new contentDetail();//각 단락 내용이 담긴 클래스

        public Content(string path)
        {
            file = path + @"\content.xml";
        }

        public void writeOdf()
        {
            StreamWriter writer = File.CreateText(file);
            string str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n" +
                "<office:document-content xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:rpt=\"http://openoffice.org/2005/report\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:ooow=\"http://openoffice.org/2004/writer\" xmlns:oooc=\"http://openoffice.org/2004/calc\" xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\" xmlns:tableooo=\"http://openoffice.org/2009/table\" xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\" xmlns:drawooo=\"http://openoffice.org/2010/draw\" xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\" xmlns:field=\"urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:dom=\"http://www.w3.org/2001/xml-events\" xmlns:xforms=\"http://www.w3.org/2002/xforms\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:formx=\"urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0\" xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:grddl=\"http://www.w3.org/2003/g/data-view#\" xmlns:css3t=\"http://www.w3.org/TR/css3-text/\" xmlns:officeooo=\"http://openoffice.org/2009/office\" office:version=\"1.3\"><office:scripts/><office:font-face-decls><style:font-face style:name=\"Lucida Sans1\" svg:font-family=\"&apos;Lucida Sans&apos;\" style:font-family-generic=\"swiss\"/><style:font-face style:name=\"바탕\" svg:font-family=\"바탕\" style:font-family-generic=\"roman\" style:font-pitch=\"variable\"/><style:font-face style:name=\"Liberation Sans\" svg:font-family=\"&apos;Liberation Sans&apos;\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\"/><style:font-face style:name=\"Lucida Sans\" svg:font-family=\"&apos;Lucida Sans&apos;\" style:font-family-generic=\"system\" style:font-pitch=\"variable\"/><style:font-face style:name=\"맑은 고딕\" svg:font-family=\"&apos;맑은 고딕&apos;\" style:font-family-generic=\"system\" style:font-pitch=\"variable\"/></office:font-face-decls><office:automatic-styles>";
            for(int i = 0; i < count; i++)
            {
                str = str + detail.textStyle();

            }
            str = str +
                "</office:automatic-styles><office:body><office:text>";
            for (int i = 0; i < count; i++)
            {
                str = str + detail.text();

            }
            str = str + "</office:text></office:body></office:document-content>";
            writer.Write(str);
            writer.Close();
        }
        public void delSubFile()
        {
            File.Delete(file);
        }
    }

    class mamifest
    {
        private string file;

        public mamifest(string path)
        {
            file = path + @"\manifest.rdf";
        }
        public void writeOdf()
        {
            
            string str = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
                "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\r\n" +
                "  <rdf:Description rdf:about=\"styles.xml\">\r\n" +
                "    <rdf:type rdf:resource=\"http://docs.oasis-open.org/ns/office/1.2/meta/odf#StylesFile\"/>\r\n" +
                "  </rdf:Description>\r\n" +
                "  <rdf:Description rdf:about=\"\">\r\n" +
                "    <ns0:hasPart xmlns:ns0=\"http://docs.oasis-open.org/ns/office/1.2/meta/pkg#\" rdf:resource=\"styles.xml\"/>\r\n" +
                "  </rdf:Description>\r\n" +
                "  <rdf:Description rdf:about=\"content.xml\">\r\n" +
                "    <rdf:type rdf:resource=\"http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile\"/>\r\n" +
                "  </rdf:Description>\r\n" +
                "  <rdf:Description rdf:about=\"\">\r\n" +
                "    <ns0:hasPart xmlns:ns0=\"http://docs.oasis-open.org/ns/office/1.2/meta/pkg#\" rdf:resource=\"content.xml\"/>\r\n" +
                "  </rdf:Description>\r\n" +
                "  <rdf:Description rdf:about=\"\">\r\n" +
                "    <rdf:type rdf:resource=\"http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document\"/>\r\n" +
                "  </rdf:Description>\r\n" +
                "</rdf:RDF>";
            StreamWriter writer = File.CreateText(file);
            writer.Write(str);
            writer.Close();
        }
        public void delSubFile()
        {
            File.Delete(file);
        }

    }
    
    class Meta
    {
        private string file;

        public Meta(string path)
        {
            file = path + @"\meta.xml";
        }
        public void writeOdf()
        {
            string str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n" +
                "<office:document-meta xmlns:grddl=\"http://www.w3.org/2003/g/data-view#\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" office:version=\"1.3\"><office:meta><meta:creation-date>2021-05-17T18:04:21.124000000</meta:creation-date><dc:date>2021-05-17T18:05:13.017000000</dc:date><meta:editing-duration>PT53S</meta:editing-duration><meta:editing-cycles>1</meta:editing-cycles><meta:document-statistic meta:table-count=\"0\" meta:image-count=\"0\" meta:object-count=\"0\" meta:page-count=\"1\" meta:paragraph-count=\"1\" meta:word-count=\"1\" meta:character-count=\"111\" meta:non-whitespace-character-count=\"111\"/><meta:generator>LibreOffice/7.1.2.2$Windows_X86_64 LibreOffice_project/8a45595d069ef5570103caea1b71cc9d82b2aae4</meta:generator></office:meta></office:document-meta>";
            StreamWriter writer = File.CreateText(file);
            writer.Write(str);
            writer.Close();
        }
        public void delSubFile()
        {
            File.Delete(file);
        }

    }

    class Mimetype
    {
        private string file;
        public Mimetype(string path)
        {
            file = path + @"\mimetype";
        }

        public void writeOdf()
        {
            string str = "application/vnd.oasis.opendocument.text";
            StreamWriter writer = File.CreateText(file);
            writer.Write(str);
            writer.Close();
        }
        public void delSubFile()
        {
            File.Delete(file);
        }
    }

    class Settings
    {
        private string file;

        public Settings(string path)
        {
            file = path + @"\settings.xml";
        }
        public void writeOdf()
        {
            string str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n" +
                "<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:config=\"urn:oasis:names:tc:opendocument:xmlns:config:1.0\" office:version=\"1.3\"><office:settings><config:config-item-set config:name=\"ooo:view-settings\"><config:config-item config:name=\"ViewAreaTop\" config:type=\"long\">0</config:config-item><config:config-item config:name=\"ViewAreaLeft\" config:type=\"long\">0</config:config-item><config:config-item config:name=\"ViewAreaWidth\" config:type=\"long\">49056</config:config-item><config:config-item config:name=\"ViewAreaHeight\" config:type=\"long\">22862</config:config-item><config:config-item config:name=\"ShowRedlineChanges\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"InBrowseMode\" config:type=\"boolean\">false</config:config-item><config:config-item-map-indexed config:name=\"Views\"><config:config-item-map-entry><config:config-item config:name=\"ViewId\" config:type=\"string\">view2</config:config-item><config:config-item config:name=\"ViewLeft\" config:type=\"long\">32364</config:config-item><config:config-item config:name=\"ViewTop\" config:type=\"long\">2988</config:config-item><config:config-item config:name=\"VisibleLeft\" config:type=\"long\">0</config:config-item><config:config-item config:name=\"VisibleTop\" config:type=\"long\">0</config:config-item><config:config-item config:name=\"VisibleRight\" config:type=\"long\">49054</config:config-item><config:config-item config:name=\"VisibleBottom\" config:type=\"long\">22860</config:config-item><config:config-item config:name=\"ZoomType\" config:type=\"short\">0</config:config-item><config:config-item config:name=\"ViewLayoutColumns\" config:type=\"short\">1</config:config-item><config:config-item config:name=\"ViewLayoutBookMode\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ZoomFactor\" config:type=\"short\">100</config:config-item><config:config-item config:name=\"IsSelectedFrame\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"AnchoredTextOverflowLegacy\" config:type=\"boolean\">false</config:config-item></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set><config:config-item-set config:name=\"ooo:configuration-settings\"><config:config-item config:name=\"ProtectForm\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrinterName\" config:type=\"string\"/><config:config-item config:name=\"EmbeddedDatabaseName\" config:type=\"string\"/><config:config-item config:name=\"CurrentDatabaseDataSource\" config:type=\"string\"/><config:config-item config:name=\"LinkUpdateMode\" config:type=\"short\">1</config:config-item><config:config-item config:name=\"AddParaTableSpacingAtStart\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"FloattableNomargins\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"UnbreakableNumberings\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"FieldAutoUpdate\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"AddVerticalFrameOffsets\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"BackgroundParaOverDrawings\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"AddParaTableSpacing\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"ChartAutoUpdate\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"CurrentDatabaseCommand\" config:type=\"string\"/><config:config-item config:name=\"PrinterSetup\" config:type=\"base64Binary\"/><config:config-item config:name=\"AlignTabStopPosition\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrinterPaperFromSetup\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"IsKernAsianPunctuation\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"CharacterCompressionType\" config:type=\"short\">0</config:config-item><config:config-item config:name=\"ApplyUserData\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"DoNotJustifyLinesWithManualBreak\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"SaveThumbnail\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"SaveGlobalDocumentLinks\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"SmallCapsPercentage66\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"CurrentDatabaseCommandType\" config:type=\"int\">0</config:config-item><config:config-item config:name=\"SaveVersionOnClose\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"UpdateFromTemplate\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"DoNotCaptureDrawObjsOnPage\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"UseFormerObjectPositioning\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintSingleJobs\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"EmbedSystemFonts\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrinterIndependentLayout\" config:type=\"string\">high-resolution</config:config-item><config:config-item config:name=\"IsLabelDocument\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"AddFrameOffsets\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"AddExternalLeading\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"MsWordCompMinLineHeightByFly\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"UseOldNumbering\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"OutlineLevelYieldsNumbering\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"DoNotResetParaAttrsForNumFont\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"IgnoreFirstLineIndentInNumbering\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"AllowPrintJobCancel\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"UseFormerLineSpacing\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"AddParaSpacingToTableCells\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"AddParaLineSpacingToTableCells\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"UseFormerTextWrapping\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"RedlineProtectionKey\" config:type=\"base64Binary\"/><config:config-item config:name=\"ConsiderTextWrapOnObjPos\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"TableRowKeep\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"TabsRelativeToIndent\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"IgnoreTabsAndBlanksForLineCalculation\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"RsidRoot\" config:type=\"int\">517029</config:config-item><config:config-item config:name=\"LoadReadonly\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ClipAsCharacterAnchoredWriterFlyFrames\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"UnxForceZeroExtLeading\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"UseOldPrinterMetrics\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"TabAtLeftIndentForParagraphsInList\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"Rsid\" config:type=\"int\">517029</config:config-item><config:config-item config:name=\"MsWordCompTrailingBlanks\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"MathBaselineAlignment\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"InvertBorderSpacing\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"CollapseEmptyCellPara\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"TabOverflow\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"StylesNoDefault\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ClippedPictures\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"EmbedFonts\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"EmbedOnlyUsedFonts\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"EmbedLatinScriptFonts\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"EmbedAsianScriptFonts\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"EmptyDbFieldHidesPara\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"EmbedComplexScriptFonts\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"TabOverMargin\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"TreatSingleColumnBreakAsPageBreak\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"SurroundTextWrapSmall\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ApplyParagraphMarkFormatToNumbering\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PropLineSpacingShrinksFirstLine\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"SubtractFlysAnchoredAtFlys\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"DisableOffPagePositioning\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ContinuousEndnotes\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ProtectBookmarks\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"ProtectFields\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"HeaderSpacingBelowLastPara\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"FrameAutowidthWithMorePara\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintAnnotationMode\" config:type=\"short\">0</config:config-item><config:config-item config:name=\"PrintGraphics\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintBlackFonts\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintLeftPages\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintControls\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintPageBackground\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintTextPlaceholder\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintDrawings\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintHiddenText\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintProspect\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintTables\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintProspectRTL\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintReversed\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintRightPages\" config:type=\"boolean\">true</config:config-item><config:config-item config:name=\"PrintFaxName\" config:type=\"string\"/><config:config-item config:name=\"PrintPaperFromSetup\" config:type=\"boolean\">false</config:config-item><config:config-item config:name=\"PrintEmptyPages\" config:type=\"boolean\">true</config:config-item></config:config-item-set></office:settings></office:document-settings>";
            StreamWriter writer = File.CreateText(file);
            writer.Write(str);
            writer.Close();
        }
        public void delSubFile()
        {
            File.Delete(file);
        }
    }

    class Styles
    {
        private string file;

        public Styles(string path)
        {
            file = path + @"\styles.xml";
        }

        public void writeOdf()
        {
            string str = "<?xml version =\"1.0\" encoding=\"UTF-8\"?>\r\n" +
                   "<office:document-styles xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:rpt=\"http://openoffice.org/2005/report\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:ooow=\"http://openoffice.org/2004/writer\" xmlns:oooc=\"http://openoffice.org/2004/calc\" xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\" xmlns:tableooo=\"http://openoffice.org/2009/table\" xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\" xmlns:drawooo=\"http://openoffice.org/2010/draw\" xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\" xmlns:field=\"urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:dom=\"http://www.w3.org/2001/xml-events\" xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:grddl=\"http://www.w3.org/2003/g/data-view#\" xmlns:css3t=\"http://www.w3.org/TR/css3-text/\" xmlns:officeooo=\"http://openoffice.org/2009/office\" office:version=\"1.3\"><office:font-face-decls><style:font-face style:name=\"Lucida Sans1\" svg:font-family=\"&apos;Lucida Sans&apos;\" style:font-family-generic=\"swiss\"/><style:font-face style:name=\"바탕\" svg:font-family=\"바탕\" style:font-family-generic=\"roman\" style:font-pitch=\"variable\"/><style:font-face style:name=\"Liberation Sans\" svg:font-family=\"&apos;Liberation Sans&apos;\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\"/><style:font-face style:name=\"Lucida Sans\" svg:font-family=\"&apos;Lucida Sans&apos;\" style:font-family-generic=\"system\" style:font-pitch=\"variable\"/><style:font-face style:name=\"맑은 고딕\" svg:font-family=\"&apos;맑은 고딕&apos;\" style:font-family-generic=\"system\" style:font-pitch=\"variable\"/></office:font-face-decls><office:styles><style:default-style style:family=\"graphic\"><style:graphic-properties svg:stroke-color=\"#3465a4\" draw:fill-color=\"#729fcf\" fo:wrap-option=\"no-wrap\" draw:shadow-offset-x=\"0.3cm\" draw:shadow-offset-y=\"0.3cm\" draw:start-line-spacing-horizontal=\"0.283cm\" draw:start-line-spacing-vertical=\"0.283cm\" draw:end-line-spacing-horizontal=\"0.283cm\" draw:end-line-spacing-vertical=\"0.283cm\" style:flow-with-text=\"false\"/><style:paragraph-properties style:text-autospace=\"ideograph-alpha\" style:line-break=\"strict\" style:writing-mode=\"lr-tb\" style:font-independent-line-spacing=\"false\"><style:tab-stops/></style:paragraph-properties><style:text-properties style:use-window-font-color=\"true\" loext:opacity=\"0%\" style:font-name=\"바탕\" fo:font-size=\"12pt\" fo:language=\"en\" fo:country=\"US\" style:letter-kerning=\"true\" style:font-name-asian=\"맑은 고딕\" style:font-size-asian=\"10.5pt\" style:language-asian=\"ko\" style:country-asian=\"KR\" style:font-name-complex=\"Lucida Sans\" style:font-size-complex=\"12pt\" style:language-complex=\"hi\" style:country-complex=\"IN\"/></style:default-style><style:default-style style:family=\"paragraph\"><style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" fo:hyphenation-ladder-count=\"no-limit\" style:text-autospace=\"ideograph-alpha\" style:punctuation-wrap=\"hanging\" style:line-break=\"strict\" style:tab-stop-distance=\"1.251cm\" style:writing-mode=\"page\"/><style:text-properties style:use-window-font-color=\"true\" loext:opacity=\"0%\" style:font-name=\"바탕\" fo:font-size=\"12pt\" fo:language=\"en\" fo:country=\"US\" style:letter-kerning=\"true\" style:font-name-asian=\"맑은 고딕\" style:font-size-asian=\"10.5pt\" style:language-asian=\"ko\" style:country-asian=\"KR\" style:font-name-complex=\"Lucida Sans\" style:font-size-complex=\"12pt\" style:language-complex=\"hi\" style:country-complex=\"IN\" fo:hyphenate=\"false\" fo:hyphenation-remain-char-count=\"2\" fo:hyphenation-push-char-count=\"2\" loext:hyphenation-no-caps=\"false\"/></style:default-style><style:default-style style:family=\"table\"><style:table-properties table:border-model=\"collapsing\"/></style:default-style><style:default-style style:family=\"table-row\"><style:table-row-properties fo:keep-together=\"auto\"/></style:default-style><style:style style:name=\"Standard\" style:family=\"paragraph\" style:class=\"text\"><style:paragraph-properties style:text-autospace=\"none\"/></style:style><style:style style:name=\"Heading\" style:family=\"paragraph\" style:parent-style-name=\"Standard\" style:next-style-name=\"Text_20_body\" style:class=\"text\"><style:paragraph-properties fo:margin-top=\"0.423cm\" fo:margin-bottom=\"0.212cm\" style:contextual-spacing=\"false\" fo:keep-with-next=\"always\"/><style:text-properties style:font-name=\"Liberation Sans\" fo:font-family=\"&apos;Liberation Sans&apos;\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\" fo:font-size=\"14pt\" style:font-name-asian=\"맑은 고딕\" style:font-family-asian=\"&apos;맑은 고딕&apos;\" style:font-family-generic-asian=\"system\" style:font-pitch-asian=\"variable\" style:font-size-asian=\"14pt\" style:font-name-complex=\"Lucida Sans\" style:font-family-complex=\"&apos;Lucida Sans&apos;\" style:font-family-generic-complex=\"system\" style:font-pitch-complex=\"variable\" style:font-size-complex=\"14pt\"/></style:style><style:style style:name=\"Text_20_body\" style:display-name=\"Text body\" style:family=\"paragraph\" style:parent-style-name=\"Standard\" style:class=\"text\"><style:paragraph-properties fo:margin-top=\"0cm\" fo:margin-bottom=\"0.247cm\" style:contextual-spacing=\"false\" fo:line-height=\"115%\"/></style:style><style:style style:name=\"List\" style:family=\"paragraph\" style:parent-style-name=\"Text_20_body\" style:class=\"list\"><style:text-properties style:font-size-asian=\"12pt\" style:font-name-complex=\"Lucida Sans1\" style:font-family-complex=\"&apos;Lucida Sans&apos;\" style:font-family-generic-complex=\"swiss\"/></style:style><style:style style:name=\"Caption\" style:family=\"paragraph\" style:parent-style-name=\"Standard\" style:class=\"extra\"><style:paragraph-properties fo:margin-top=\"0.212cm\" fo:margin-bottom=\"0.212cm\" style:contextual-spacing=\"false\" text:number-lines=\"false\" text:line-number=\"0\"/><style:text-properties fo:font-size=\"12pt\" fo:font-style=\"italic\" style:font-size-asian=\"12pt\" style:font-style-asian=\"italic\" style:font-name-complex=\"Lucida Sans1\" style:font-family-complex=\"&apos;Lucida Sans&apos;\" style:font-family-generic-complex=\"swiss\" style:font-size-complex=\"12pt\" style:font-style-complex=\"italic\"/></style:style><style:style style:name=\"Index\" style:family=\"paragraph\" style:parent-style-name=\"Standard\" style:class=\"index\"><style:paragraph-properties text:number-lines=\"false\" text:line-number=\"0\"/><style:text-properties style:font-size-asian=\"12pt\" style:font-name-complex=\"Lucida Sans1\" style:font-family-complex=\"&apos;Lucida Sans&apos;\" style:font-family-generic-complex=\"swiss\"/></style:style><text:outline-style style:name=\"Outline\"><text:outline-level-style text:level=\"1\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"2\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"3\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"4\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"5\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"6\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"7\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"8\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"9\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style><text:outline-level-style text:level=\"10\" style:num-format=\"\"><style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment text:label-followed-by=\"listtab\"/></style:list-level-properties></text:outline-level-style></text:outline-style><text:notes-configuration text:note-class=\"footnote\" style:num-format=\"1\" text:start-value=\"0\" text:footnotes-position=\"page\" text:start-numbering-at=\"document\"/><text:notes-configuration text:note-class=\"endnote\" style:num-format=\"i\" text:start-value=\"0\"/><text:linenumbering-configuration text:number-lines=\"false\" text:offset=\"0.499cm\" style:num-format=\"1\" text:number-position=\"left\" text:increment=\"5\"/></office:styles><office:automatic-styles><style:page-layout style:name=\"Mpm1\"><style:page-layout-properties fo:page-width=\"21.001cm\" fo:page-height=\"29.7cm\" style:num-format=\"1\" style:print-orientation=\"portrait\" fo:margin-top=\"2cm\" fo:margin-bottom=\"2cm\" fo:margin-left=\"2cm\" fo:margin-right=\"2cm\" style:writing-mode=\"lr-tb\" style:footnote-max-height=\"0cm\"><style:footnote-sep style:width=\"0.018cm\" style:distance-before-sep=\"0.101cm\" style:distance-after-sep=\"0.101cm\" style:line-style=\"solid\" style:adjustment=\"left\" style:rel-width=\"25%\" style:color=\"#000000\"/></style:page-layout-properties><style:header-style/><style:footer-style/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name=\"Standard\" style:page-layout-name=\"Mpm1\"/></office:master-styles></office:document-styles>";
            StreamWriter writer = File.CreateText(file);
            writer.Write(str);
            writer.Close();
        }
        public void delSubFile()
        {
            File.Delete(file);
        }

    }
    

}
