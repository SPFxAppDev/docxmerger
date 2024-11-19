import JSZip, { JSZipObject } from 'jszip';

export type DocxMergerOptions = {
  /**
   * Determines whether a page break should be inserted after merging the file(s)
   * @default true
   * @type {boolean}
   * @since 1.0.0
   */
  pageBreak?: boolean;

  /**
   * Optional parameters that can be passed to JSZip and the loadAsync method when the merge method is called
   * @default undefined
   * @type {JSZip.JSZipLoadOptions}
   * @since 1.0.0
   */
  jsZipLoadOptions?: JSZip.JSZipLoadOptions;

  /**
   * Optional parameters that can be passed to JSZip and the generateAsync method when the save method is called
   * @default { type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: {level: 4} }
   * @type {JSZip.JSZipGeneratorOptions}
   * @since 1.0.0
   */
  jsZipGenerateOptions?: JSZip.JSZipGeneratorOptions;
};

/**
 * A type of file that can be used for merging
 * @since 1.0.0
 */
export type JSZipFileInput = ArrayBuffer | Uint8Array | Blob;

/**
 *
 * The interface that the DocxMerger implements
 * @export
 * @interface IDocxMerger
 */
export interface IDocxMerger {
  /**
   * Merges the passed files into a single file
   * @param {JSZipFileInput[]} files the files to merge
   * @param {DocxMergerOptions} [options] Optional options for the merging and saving process
   */
  merge(files: JSZipFileInput[], options?: DocxMergerOptions): Promise<void>;

  /**
   * Creates/saves the merged file and returns it
   * @returns a Promise that resolves to any type of value.
   */
  save(): Promise<any>;
}

export class DocxMerger implements IDocxMerger {
  private body: any = [];
  // private header = [];
  // private footer = [];
  //   private pageBreak = true;
  // private Basestyle = 'source';
  private style: any = [];
  private numbering: any = [];
  private files: JSZip[] = [];
  private contentTypes: any = [];
  private media: any = {};
  private rel: any = {};
  private builder: string[] = [];
  private options: DocxMergerOptions = {};
  private mediaFilesCount: number = 1;

  private contentTypeAndValueMapper: Record<string, Record<string, any>> = {};

  public constructor() {
    this.builder = this.body;
  }

  /**
   * Merges the passed files into a single file
   * @param {JSZipFileInput[]} files the files to merge
   * @param {DocxMergerOptions} [options] Optional options for the merging and saving process
   * @memberof DocxMerger
   */
  public async merge(files: JSZipFileInput[], options?: DocxMergerOptions): Promise<void> {
    this.mediaFilesCount = 1;
    this.contentTypeAndValueMapper = {};

    const defaultOptions: DocxMergerOptions = {
      pageBreak: true,
      jsZipGenerateOptions: {
        type: 'arraybuffer',
        compression: 'DEFLATE',
        compressionOptions: {
          level: 4,
        },
      },
    };

    this.options = { ...defaultOptions, ...options };

    files = files || [];
    // this.Basestyle = options.style || 'source';

    for (const file of files) {
      const zipFile: JSZip = await new JSZip().loadAsync(file, this.options.jsZipLoadOptions);
      this.files.push(zipFile);
    }

    if (this.files.length > 0) {
      await this.mergeBody();
    }
  }

  /**
   * Creates/saves the merged file and returns it
   * @returns a Promise that resolves to any type of value.
   */
  public async save(): Promise<any> {
    if (!this.files || (this.files && this.files.length < 1)) {
      return;
    }

    const zip: JSZip = this.files[0] as JSZip;

    let xmlString = await (zip.file('word/document.xml') as JSZip.JSZipObject).async('string');

    const startIndex = xmlString.indexOf('<w:body>') + 8;
    const endIndex = xmlString.lastIndexOf('<w:sectPr');

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), this.body.join(''));

    await this.generateContentTypes(zip);
    await this.copyMediaFiles(zip);
    await this.generateRelations(zip);
    await this.generateNumbering(zip);
    await this.generateStyles(zip);

    zip.file('word/document.xml', xmlString);

    this.options.jsZipGenerateOptions = this.options.jsZipGenerateOptions||{};
    this.options.jsZipGenerateOptions.type =
      this.options.jsZipGenerateOptions.type || 'arraybuffer';
    return await zip.generateAsync(this.options.jsZipGenerateOptions);
  }

  private async mergeBody(): Promise<void> {
    this.builder = this.body;

    const numberSerializer = new XMLSerializer();
    const styleSerializer = new XMLSerializer();

    for (let index: number = 0; index < this.files.length; index++) {
      const zip: JSZip = this.files[index];
      await this.mergeContentTypes(zip);
      await this.prepareMediaFiles(zip, index);
      await this.mergeRelations(zip);

      await this.prepareNumbering(zip, index, numberSerializer);
      await this.mergeNumbering(zip);
      await this.prepareStyles(zip, index, styleSerializer);
      await this.mergeStyles(zip);

      let xmlString = await (zip.file('word/document.xml') as JSZip.JSZipObject).async('string');
      xmlString = xmlString.substring(xmlString.indexOf('<w:body>') + 8);
      xmlString = xmlString.substring(0, xmlString.indexOf('</w:body>'));
      xmlString = xmlString.substring(0, xmlString.lastIndexOf('<w:sectPr'));

      this.insertRaw(xmlString);

      if (this.options.pageBreak && index < this.files.length - 1) {
        this.insertPageBreak();
      }
    }
  }

  private insertPageBreak(): void {
    const pb =
      '<w:p> \
                        <w:r> \
                            <w:br w:type="page"/> \
                        </w:r> \
                      </w:p>';

    this.builder.push(pb);
  }

  private insertRaw(xml: string): void {
    this.builder.push(xml);
  }

  //#region ContentTypes
  private async mergeContentTypes(zipFile: JSZip): Promise<void> {
    const xmlString = await (zipFile.file('[Content_Types].xml') as JSZip.JSZipObject).async(
      'string'
    );
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    const childNodes: any = xml.getElementsByTagName('Types')[0].childNodes;

    for (const childNode in childNodes) {
      const node = childNode as any;
      if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
        const contentType = childNodes[node].getAttribute('ContentType');
        let value =
          childNodes[node].getAttribute('PartName') || childNodes[node].getAttribute('Extension');

        let addToCollection = false;
        if (!this.contentTypeAndValueMapper[contentType]) {
          this.contentTypeAndValueMapper[contentType] = {};
        }

        let key = value || 'EMPTY';

        if (!this.contentTypeAndValueMapper[contentType][key]) {
          this.contentTypeAndValueMapper[contentType][key] = {};
          addToCollection = true;
        }

        if (addToCollection) {
          this.contentTypes.push(childNodes[node].cloneNode());
        }
      }
    }
  }

  private async generateContentTypes(zipFile: JSZip): Promise<void> {
    let xmlString = await (zipFile.file('[Content_Types].xml') as JSZip.JSZipObject).async(
      'string'
    );
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const serializer = new XMLSerializer();

    const types = xml.documentElement.cloneNode();

    for (const node in this.contentTypes) {
      types.appendChild(this.contentTypes[node]);
    }

    const startIndex = xmlString.indexOf('<Types');
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zipFile.file('[Content_Types].xml', xmlString);
  }
  //#endregion ContentTypes

  //#region Relations
  private async mergeRelations(zipFile: JSZip): Promise<void> {
    const xmlString = await (
      zipFile.file('word/_rels/document.xml.rels') as JSZip.JSZipObject
    ).async('string');
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    const childNodes: any = xml.getElementsByTagName('Relationships')[0].childNodes;

    for (const node in childNodes) {
      if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
        const Id = childNodes[node].getAttribute('Id');
        if (!this.rel[Id]) {
          this.rel[Id] = childNodes[node].cloneNode();
        }
      }
    }
  }

  private async generateRelations(zipFile: JSZip): Promise<void> {
    let xmlString = await (zipFile.file('word/_rels/document.xml.rels') as JSZip.JSZipObject).async(
      'string'
    );
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const serializer = new XMLSerializer();

    const types = xml.documentElement.cloneNode();

    for (const node in this.rel) {
      types.appendChild(this.rel[node]);
    }

    const startIndex = xmlString.indexOf('<Relationships');
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zipFile.file('word/_rels/document.xml.rels', xmlString);
  }
  //#endregion Relations

  //#region MEDIA FILES
  private async prepareMediaFiles(zipFile: JSZip, mergedFileIndex: number): Promise<void> {
    const medFiles = (zipFile.folder('word/media') as JSZip).files;
    for (const mfile in medFiles) {
      if (/^word\/media/.test(mfile) && mfile.length > 11) {
        this.media[this.mediaFilesCount] = {};
        this.media[this.mediaFilesCount].oldTarget = mfile;
        this.media[this.mediaFilesCount].newTarget = mfile
          .replace(/[0-9]/, '_' + this.mediaFilesCount)
          .replace('word/', '');
        this.media[this.mediaFilesCount].fileIndex = mergedFileIndex;
        await this.updateMediaRelations(zipFile);
        await this.updateMediaContent(zipFile);
        this.mediaFilesCount++;
      }
    }
  }

  private async updateMediaRelations(zipFile: JSZip): Promise<void> {
    let xmlString = await (zipFile.file('word/_rels/document.xml.rels') as JSZip.JSZipObject).async(
      'string'
    );
    let xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    const childNodes: any = xml.getElementsByTagName('Relationships')[0].childNodes;
    const serializer = new XMLSerializer();

    for (const node in childNodes) {
      if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
        const target = childNodes[node].getAttribute('Target');
        if ('word/' + target === this.media[this.mediaFilesCount].oldTarget) {
          this.media[this.mediaFilesCount].oldRelID = childNodes[node].getAttribute('Id');

          childNodes[node].setAttribute('Target', this.media[this.mediaFilesCount].newTarget);
          childNodes[node].setAttribute(
            'Id',
            this.media[this.mediaFilesCount].oldRelID + '_' + this.mediaFilesCount
          );
        }
      }
    }

    const startIndex = xmlString.indexOf('<Relationships');
    xmlString = xmlString.replace(
      xmlString.slice(startIndex),
      serializer.serializeToString(xml.documentElement)
    );

    zipFile.file('word/_rels/document.xml.rels', xmlString);
  }

  private async updateMediaContent(zipFile: JSZip): Promise<void> {
    let xmlString = await (zipFile.file('word/document.xml') as JSZip.JSZipObject).async('string');
    xmlString = xmlString.replace(
      new RegExp(this.media[this.mediaFilesCount].oldRelID + '"', 'g'),
      this.media[this.mediaFilesCount].oldRelID + '_' + this.mediaFilesCount + '"'
    );
    zipFile.file('word/document.xml', xmlString);
  }

  private async copyMediaFiles(zipFile: JSZip): Promise<void> {
    for (const media in this.media) {
      const fileIndex: number = this.media[media].fileIndex;
      const content = await (
        (this.files[fileIndex] as JSZip).file(this.media[media].oldTarget) as JSZipObject
      ).async('uint8array');
      zipFile.file('word/' + this.media[media].newTarget, content);
    }
  }
  //#endregion MEDIA FILES

  //#region Numbering
  private async prepareNumbering(
    zipFile: JSZip,
    mergedFileIndex: number,
    serializer: XMLSerializer
  ): Promise<void> {
    const xmlBin = zipFile.file('word/numbering.xml');
    if (!xmlBin) {
      return;
    }
    let xmlString = await xmlBin.async('string');
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const nodes: any = xml.getElementsByTagName('w:abstractNum');

    for (const node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        const absID = nodes[node].getAttribute('w:abstractNumId');
        nodes[node].setAttribute('w:abstractNumId', absID + mergedFileIndex);
        const pStyles = nodes[node].getElementsByTagName('w:pStyle');
        for (const pStyle in pStyles) {
          if (pStyles[pStyle].getAttribute) {
            const pStyleId = pStyles[pStyle].getAttribute('w:val');
            pStyles[pStyle].setAttribute('w:val', pStyleId + '_' + mergedFileIndex);
          }
        }
        const numStyleLinks = nodes[node].getElementsByTagName('w:numStyleLink');
        for (const numstyleLink in numStyleLinks) {
          if (numStyleLinks[numstyleLink].getAttribute) {
            const styleLinkId = numStyleLinks[numstyleLink].getAttribute('w:val');
            numStyleLinks[numstyleLink].setAttribute('w:val', styleLinkId + '_' + mergedFileIndex);
          }
        }

        const styleLinks = nodes[node].getElementsByTagName('w:styleLink');
        for (const styleLink in styleLinks) {
          if (styleLinks[styleLink].getAttribute) {
            const styleLinkId = styleLinks[styleLink].getAttribute('w:val');
            styleLinks[styleLink].setAttribute('w:val', styleLinkId + '_' + mergedFileIndex);
          }
        }
      }
    }

    const numNodes: any = xml.getElementsByTagName('w:num');

    for (const node in numNodes) {
      if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
        const ID = numNodes[node].getAttribute('w:numId');
        numNodes[node].setAttribute('w:numId', ID + mergedFileIndex);
        const absrefID = numNodes[node].getElementsByTagName('w:abstractNumId');
        for (const i in absrefID) {
          if (absrefID[i].getAttribute) {
            const iId = absrefID[i].getAttribute('w:val');
            absrefID[i].setAttribute('w:val', iId + mergedFileIndex);
          }
        }
      }
    }

    const startIndex = xmlString.indexOf('<w:numbering ');
    xmlString = xmlString.replace(
      xmlString.slice(startIndex),
      serializer.serializeToString(xml.documentElement)
    );

    zipFile.file('word/numbering.xml', xmlString);
  }

  private async mergeNumbering(zipFile: JSZip): Promise<void> {
    const xmlBin = zipFile.file('word/numbering.xml');
    if (!xmlBin) {
      return;
    }
    let xmlString = await xmlBin.async('string');
    xmlString = xmlString.substring(
      xmlString.indexOf('<w:abstractNum '),
      xmlString.indexOf('</w:numbering')
    );
    this.numbering.push(xmlString);
  }

  private async generateNumbering(zipFile: JSZip): Promise<void> {
    const xmlBin = zipFile.file('word/numbering.xml');
    if (!xmlBin) {
      return;
    }
    let xmlString = await xmlBin.async('string');
    const startIndex = xmlString.indexOf('<w:abstractNum ');
    const endIndex = xmlString.indexOf('</w:numbering>');

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), this.numbering.join(''));

    zipFile.file('word/numbering.xml', xmlString);
  }
  //#endregion Numbering

  //#region STYLING
  private async prepareStyles(
    zipFile: JSZip,
    mergedFileIndex: number,
    serializer: XMLSerializer
  ): Promise<void> {
    let xmlString = await (zipFile.file('word/styles.xml') as JSZip.JSZipObject).async('string');
    let xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const nodes: any = xml.getElementsByTagName('w:style');

    for (const node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        const styleId = nodes[node].getAttribute('w:styleId');
        nodes[node].setAttribute('w:styleId', styleId + '_' + mergedFileIndex);
        const basedonStyle = nodes[node].getElementsByTagName('w:basedOn')[0];
        if (basedonStyle) {
          const basedonStyleId = basedonStyle.getAttribute('w:val');
          basedonStyle.setAttribute('w:val', basedonStyleId + '_' + mergedFileIndex);
        }

        const w_next = nodes[node].getElementsByTagName('w:next')[0];
        if (w_next) {
          const w_next_ID = w_next.getAttribute('w:val');
          w_next.setAttribute('w:val', w_next_ID + '_' + mergedFileIndex);
        }

        const w_link = nodes[node].getElementsByTagName('w:link')[0];
        if (w_link) {
          const w_link_ID = w_link.getAttribute('w:val');
          w_link.setAttribute('w:val', w_link_ID + '_' + mergedFileIndex);
        }

        const numId = nodes[node].getElementsByTagName('w:numId')[0];
        if (numId) {
          const numId_ID = numId.getAttribute('w:val');
          numId.setAttribute('w:val', numId_ID + mergedFileIndex);
        }

        await this.updateStyleRel_Content(zipFile, mergedFileIndex, styleId);
      }
    }

    const startIndex = xmlString.indexOf('<w:styles ');
    xmlString = xmlString.replace(
      xmlString.slice(startIndex),
      serializer.serializeToString(xml.documentElement)
    );

    zipFile.file('word/styles.xml', xmlString);
  }

  private async updateStyleRel_Content(
    zipFile: JSZip,
    mergedFileIndex: number,
    styleId: string
  ): Promise<void> {
    let xmlString = await (zipFile.file('word/document.xml') as JSZip.JSZipObject).async('string');
    xmlString = xmlString.replace(
      new RegExp('w:val="' + styleId + '"', 'g'),
      'w:val="' + styleId + '_' + mergedFileIndex + '"'
    );
    zipFile.file('word/document.xml', xmlString);
  }

  private async mergeStyles(zipFile: JSZip): Promise<void> {
    let xmlString = await (zipFile.file('word/styles.xml') as JSZip.JSZipObject).async('string');
    xmlString = xmlString.substring(
      xmlString.indexOf('<w:style '),
      xmlString.indexOf('</w:styles')
    );
    this.style.push(xmlString);
  }

  private async generateStyles(zipFile: JSZip): Promise<void> {
    let xmlString = await (zipFile.file('word/styles.xml') as JSZip.JSZipObject).async('string');
    const startIndex = xmlString.indexOf('<w:style ');
    const endIndex = xmlString.indexOf('</w:styles>');

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), this.style.join(''));

    zipFile.file('word/styles.xml', xmlString);
  }
  //#endregion STYLING
}
