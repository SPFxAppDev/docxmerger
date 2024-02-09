# @spfxappdev/docxmerger

This is a merge/copy of the [docx-merger package](https://github.com/apurvaojas/docx-merger). But since the package has not been updated for more than 5 years and the package has some bugs regarding "corrupt" Word files, I created my own package. This package also has TypeScript support and the API has been changed a bit. The JSZip npm package has also been updated. This package (unpacked) is about 1.5MB smaller than the original `docx-merger package`

## Installation

```bash
npm install @spfxappdev/docxmerger
```

## Usage

The example shows how an input field of type "File" is handled after the selection has been changed. The input field allows multiple selections:

### HTML

```HTML
<input type="file" id="wordFileInput" multiple accept=".docx" />
```

### TypeScript 

```TypeScript
import { DocxMerger } from '@spfxappdev/docxmerger';

const fileInput = document.getElementById('wordFileInput') as HTMLInputElement;

fileInput.addEventListener('change', async (event) => {
  const wordFiles = fileInput.files as FileList;
  
  if(wordFiles.length < 2) {
    alert("Please select at least 2 files");
    return;
  }
  
  //All loaded files in an array of ArrayBuffer
  const promises: Promise<ArrayBuffer>[] = [];

  //Load all files as arrayBuffer and then resolve the "Promise"
  const readFilesPromise = new Promise<void>((res, rej) => {
    for (let fileIndex = 0; fileIndex < (wordFiles as FileList).length; fileIndex++) {
      const reader = new FileReader();
      const wordFile = wordFiles[fileIndex];

      reader.onload = function (event) {
        const arrayBuffer = event.target.result;

        promises.push(Promise.resolve<ArrayBuffer>(arrayBuffer as ArrayBuffer));

        if (promises.length === wordFiles.length) {
          res();
        }
      };

      reader.readAsArrayBuffer(wordFile);
    }
  });

  await readFilesPromise;
  
  const filesToMerge = await Promise.all(promises);

  const docx = new DocxMerger();
  await docx.merge(filesToMerge);
  const mergedFile  = await docx.save();
  
  //DO something with the merged file
});
```

## Demo

In this [codepen](https://codepen.io/SimpleBase/pen/YzgOpvj) you will find an example implementation including the code on how to download the merged file after merging


## API

### merge(files: JSZipFileInput[], options?: DocxMergerOptions)
![since @spfxappdev/docxmerger@1.0.0](https://img.shields.io/badge/since-v1.0.0-orange)

This method merges the passed files into a single file.

#### Arguments

| name | type | description |
|-------|---------|-------------------------------------|
| files  | `JSZipFileInput` ( ==>`ArrayBuffer` or `Uint8Array` or `Blob`)    | the files to merge |
| options  | `DocxMergerOptions` | `Optional` options for the merging and saving process |

#### Type `DocxMergerOptions`



| name | type | description |
|-------|---------|-------------------------------------|
| pageBreak  | `boolean` | Determines whether a page break should be inserted after merging the file(s) |
| jsZipLoadOptions  | `JSZip.JSZipLoadOptions` | `Optional` parameters that can be passed to JSZip and the loadAsync method when the merge method is called [see JSZip](https://stuk.github.io/jszip/documentation/api_jszip/load_async.html) |
| jsZipGenerateOptions  | `JSZip.JSZipGeneratorOptions` | `Optional` parameters that can be passed to JSZip and the generateAsync method when the save method is called [see JSZip](https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html) <br/><br/> Default value: <br/> `{ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: {level: 4} }` |

______


### save()

![since @spfxappdev/docxmerger@1.0.0](https://img.shields.io/badge/since-v1.0.0-orange)

Creates/saves the merged file async and returns it

