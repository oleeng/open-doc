class OpenSheet {
  constructor(name) {
    this._data = {
      sheets: []
    }
  }

  async export(){
    return new Promise((resolve, reject) => {
      let zip = new JSZip();

      let content = new XMLContainer()

      zip.file("content.xml", content.export());
      /*done*/zip.file("styles.xml",  styleCreate().export());
      /*done*/zip.file("mimetype", "application/vnd.oasis.opendocument.spreadsheet");
      /*done*/zip.file("meta.xml", metaCreate().export());
      /*done*/zip.file("META-INF/manifest.xml", manifestCreate().export());

      zip.generateAsync({
        type:"blob",
        mimeType: "application/vnd.oasis.opendocument.spreadsheet",
        compression: "DEFLATE",
        compressionOptions: {
          level: 6
        }
      }).then(function(file) {
        resolve(file)
      });
    })

    function metaCreate() {
      const date =  new Date();

      let metaRoot = new XMLContainer()

      let meta = metaRoot.addChild(new Node("office:document-meta"))

      meta.addAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
      meta.addAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
      meta.addAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
      meta.addAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
      meta.addAttribute("office:version", "1.2")

      let a = meta.addChild(new Node("office:meta"))
      a.addChild(new Node("meta:generator")).addChild("MicrosoftOffice/16.0 MicrosoftExcel/CalculationVersion-23029")
      a.addChild(new Node("meta:initial-creator")).addChild("Open Doc")
      a.addChild(new Node("dc:creator")).addChild("Open Doc")
      a.addChild(new Node("meta:creation-date")).addChild(date.toISOString())
      a.addChild(new Node("dc:date")).addChild(date.toISOString())
      a.addChild(new Node("meta:editing-duration")).addChild("PT0S")

      return metaRoot
    }

    function manifestCreate() {
      let manifestRoot = new XMLContainer()

      let manifest = manifestRoot.addChild(new Node("manifest:manifest"))
      manifest.addAttribute("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")
      let a = manifest.addChild(new Node("manifest:file-entry"))
      a.addAttribute("manifest:full-path", "/")
      a.addAttribute("manifest:media-type", "application/vnd.oasis.opendocument.spreadsheet")

      a = manifest.addChild(new Node("manifest:file-entry"))
      a.addAttribute("manifest:full-path", "styles.xml")
      a.addAttribute("manifest:media-type", "text/xml")

      a = manifest.addChild(new Node("manifest:file-entry"))
      a.addAttribute("manifest:full-path", "content.xml")
      a.addAttribute("manifest:media-type", "text/xml")

      a = manifest.addChild(new Node("manifest:file-entry"))
      a.addAttribute("manifest:full-path", "meta.xml")
      a.addAttribute("manifest:media-type", "text/xml")

      return manifestRoot
    }

    function styleCreate(){
      let styleRoot = new XMLContainer()

      let style = styleRoot.addChild(new Node("office:document-styles"))
      style.addAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
      style.addAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
      style.addAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
      style.addAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
      style.addAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
      style.addAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
      style.addAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
      style.addAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
      style.addAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
      style.addAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
      style.addAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
      style.addAttribute("office:version", "1.2")

      let a = style.addChild(new Node("office:font-face-decls"))
      let b = a.addChild(new Node("style:font-face"))
      b.addAttribute("style:name", "Arial")
      b.addAttribute("svg:font-family", "Arial")

      a = style.addChild(new Node("office:styles"))

      b = a.addChild(new Node("number:number-style"))
      b.addAttribute("style:name", "N0")
      let c = b.addChild(new Node("number:number"))
      c.addAttribute("number:min-integer-digits", "1")

      b = a.addChild(new Node("number:number-style"))
      b.addAttribute("style:name", "N1")
      c = b.addChild(new Node("number:number"))
      c.addAttribute("number:decimal-places", "0")
      c.addAttribute("number:min-integer-digits", "1")

      b = a.addChild(new Node("number:text-style"))
      b.addAttribute("style:name", "N30")
      c = b.addChild(new Node("number:text-content"))

      b = a.addChild(new Node("number:currency-style"))
      b.addAttribute("style:name", "N36")
      b.addAttribute("number:language", "de")
      b.addAttribute("number:country", "DE")
      c = b.addChild(new Node("number:number"))
      c.addAttribute("number:decimal-places", "2")
      c.addAttribute("number:min-integer-digits", "1")
      c.addAttribute("number:grouping", "true")
      c = b.addChild(new Node("number:text"))
      c = b.addChild(new Node("number:currency-symbol"))
      c.addAttribute("number:language", "de")
      c.addAttribute("number:country", "DE")
      c.addChild("EUR")

      b = a.addChild(new Node("style:style"))
      b.addAttribute("style:name", "Default")
      b.addAttribute("style:family", "table-cell")
      b.addAttribute("style:data-style-name", "N0")

      c = b.addChild(new Node("style:table-cell-properties"))
      c.addAttribute("style:vertical-align", "automatic")
      c.addAttribute("fo:background-color", "transparent")

      c = b.addChild("style:text-properties")
      c.addAttribute("fo:color", "#000000")
      c.addAttribute("fo:font-name", "Arial")
      c.addAttribute("fo:font-name-asian", "Arial")
      c.addAttribute("fo:font-name-complex", "Arial")
      c.addAttribute("fo:font-size", "10pt")
      c.addAttribute("fo:font-size-asian", "10pt")
      c.addAttribute("fo:font-size-complex", "10pt")

      a = style.addChild(new Node("office:automatic-styles"))
      b = a.addChild(new Node("style:page-layout"))
      b.addAttribute("style:name", "pm1")
      c = b.addChild(new Node("style:page-layout-properties"))
      c.addAttribute("fo:margin-top", "0.3in")
      c.addAttribute("fo:margin-bottom", "0.3in")
      c.addAttribute("fo:margin-left", "0.7in")
      c.addAttribute("fo:margin-right", "0.7in")
      c.addAttribute("style:print-orientation", "landscape")
      c.addAttribute("style:print-page-order", "ttb")
      c.addAttribute("style:first-page-number", "continue")
      c.addAttribute("style:scale-to", "67%")
      c.addAttribute("style:table-centering", "none")
      c.addAttribute("style:print", "grid objects charts drawings")
      c = b.addChild(new Node("style:header-style"))
      let d = c.addChild(new Node("style:header-footer-properties"))
      d.addAttribute("fo:min-height", "0.487401575in")
      d.addAttribute("fo:margin-left", "0.7in")
      d.addAttribute("fo:margin-right", "0.7in")
      d.addAttribute("fo:margin-bottom", "0in")
      c = b.addChild(new Node("style:footer-style"))
      d = c.addChild(new Node("style:header-footer-properties"))
      d.addAttribute("fo:min-height", "0.487401575in")
      d.addAttribute("fo:margin-left", "0.7in")
      d.addAttribute("fo:margin-right", "0.7in")
      d.addAttribute("fo:margin-top", "0in")

      a = style.addChild(new Node("office:master-styles"))
      b = a.addChild(new Node("style:master-page"))
      b.addAttribute("style:name", "mp1")
      b.addAttribute("style:page-layout-name", "pm1")
      c = b.addChild(new Node("style:header"))
      c = b.addChild(new Node("style:header-left"))
      c.addAttribute("style:display", "false")
      c = b.addChild(new Node("style:footer"))
      c = b.addChild(new Node("style:footer-left"))
      c.addAttribute("style:display", "false")

      return styleRoot
    }

    function contentCreate() {

    }
  }
}
