#!/usr/bin/env node

// @ts-check

const concat        = require('concat-stream');
const { DOMParser } = require('@xmldom/xmldom');
const fileType      = require('file-type');
const fs            = require('fs');
const yauzl         = require('yauzl');

/** Load pdfjs-dist once at module scope. This returns a Promise that resolves to the module. */
const pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs');

/** Header for error messages */
const ERRORHEADER = "[OfficeParser]: ";
/** Error messages */
const ERRORMSG = {
    extensionUnsupported: (ext) =>      `Sorry, OfficeParser currently support docx, pptx, xlsx, odt, odp, ods, pdf files only. Create a ticket in Issues on github to add support for ${ext} files. Stay tuned for further updates.`,
    fileCorrupted:        (filepath) => `Your file ${filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error.`,
    fileDoesNotExist:     (filepath) => `File ${filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location.`,
    locationNotFound:     (location) => `Entered location ${location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter.`,
    improperArguments:                  `Improper arguments`,
    improperBuffers:                    `Error occured while reading the file buffers`,
    invalidInput:                       `Invalid input type: Expected a Buffer or a valid file path`
}

/** Returns parsed xml document for a given xml text.
 * @param {string} xml The xml string from the doc file
 * @returns {XMLDocument}
 */
const parseString = (xml) => {
    let parser = new DOMParser();
    return parser.parseFromString(xml, "text/xml");
};

/** @typedef {Object} OfficeParserConfig
 * @property {boolean} [outputErrorToConsole] Flag to show all the logs to console in case of an error irrespective of your own handling. Default is false.
 * @property {string}  [newlineDelimiter]     The delimiter used for every new line in places that allow multiline text like word. Default is \n.
 * @property {boolean} [ignoreNotes]          Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.
 * @property {boolean} [putNotesAtLast]       Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored.
 */


/** Main function for parsing text from word files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */

function parseWord(file, callback, config) {
    /** The target content xml file for the docx file. */
    const mainContentFileRegex = /word\/document[\d+]?.xml/g;
    const footnotesFileRegex   = /word\/footnotes[\d+]?.xml/g;
    const endnotesFileRegex    = /word\/endnotes[\d+]?.xml/g;
    const stylesFileRegex      = /word\/styles.xml/g;

    extractFiles(file, x => [mainContentFileRegex, footnotesFileRegex, endnotesFileRegex, stylesFileRegex].some(fileRegex => x.match(fileRegex)))
        .then(files => {
            // Verify if atleast the document xml file exists in the extracted files list.
            if (!files.some(file => file.path.match(mainContentFileRegex)))
                throw ERRORMSG.fileCorrupted(file);

            const stylesFile = files.find(file => file.path.match(stylesFileRegex));
            const styleMap = stylesFile ? parseStyles(stylesFile.content) : {};

            return {
                contentFiles: files
                    .filter(file => file.path.match(mainContentFileRegex) || file.path.match(footnotesFileRegex) || file.path.match(endnotesFileRegex))
                    .map(file => file.content),
                styleMap: styleMap
            };
        })
        .then(({contentFiles, styleMap}) => {
            /** Store all the text content to respond. */
            let responseText = [];

            contentFiles.forEach(xmlContent => {
                const doc = parseString(xmlContent);
                const bodyElements = doc.getElementsByTagName("w:body")[0];
                if (!bodyElements) return;

                // 获取body下的所有直接子元素
                const childElements = Array.from(bodyElements.childNodes).filter(node => node.nodeType === 1);
                
                childElements.forEach(node => {
                    /** @type {Element} */
                    // @ts-ignore
                    var element = node
                    if (element.nodeName === "w:tbl") {
                        // 处理表格
                        const markdownTable = parseWordTable(element, styleMap);
                        if (markdownTable) {
                            responseText.push(markdownTable);
                        }
                    } else if (element.nodeName === "w:p") {
                        // 处理段落
                        if (element.getElementsByTagName("w:t").length > 0) {
                            const pStyle = getParagraphStyle(element, styleMap);
                            
                            const xmlTextNodeList = element.getElementsByTagName("w:t");
                            const formattedText = Array.from(xmlTextNodeList)
                                .filter(textNode => textNode.childNodes[0] && textNode.childNodes[0].nodeValue)
                                .map(textNode => {
                                    const text = textNode.childNodes[0].nodeValue;
                                    const runNode = textNode.parentNode;
                                    const formatting = runNode ? getTextFormatting(runNode) : {};
                                    return applyMarkdownFormatting(text, formatting);
                                })
                                .join("");

                            const paragraphText = applyParagraphFormatting(formattedText, pStyle);
                            if (paragraphText.trim()) {
                                responseText.push(paragraphText);
                            }
                        }
                    }
                });
            });

            // Respond by calling the Callback function.
            callback(responseText.join(config.newlineDelimiter ?? "\n"), undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Parse Word table and convert to Markdown table
 * @param  tableElement The w:tbl element
 * @param {Object} styleMap Style mapping object
 * @returns {string} Markdown formatted table
 */
function parseWordTable(tableElement, styleMap) {
    const rows = tableElement.getElementsByTagName("w:tr");
    if (rows.length === 0) return "";

    const tableData = [];
    let maxCols = 0;

    // 解析所有行
    Array.from(rows).forEach(row => {
        const cells = row.getElementsByTagName("w:tc");
        const rowData = [];
        
        Array.from(cells).forEach(cell => {
            // 获取单元格中的所有段落
            const paragraphs = cell.getElementsByTagName("w:p");
            const cellContent = [];
            
            Array.from(paragraphs).forEach(paragraph => {
                const textNodes = paragraph.getElementsByTagName("w:t");
                if (textNodes.length > 0) {
                    const paragraphText = Array.from(textNodes)
                        .filter(textNode => textNode.childNodes[0] && textNode.childNodes[0].nodeValue)
                        .map(textNode => {
                            const text = textNode.childNodes[0].nodeValue;
                            const runNode = textNode.parentNode;
                            const formatting = runNode ? getTextFormatting(runNode) : {};
                            return applyMarkdownFormatting(text, formatting);
                        })
                        .join("");
                    
                    if (paragraphText.trim()) {
                        cellContent.push(paragraphText.trim());
                    }
                }
            });
            
            // 将单元格内容用空格连接，如果为空则用空字符串
            rowData.push(cellContent.join(" ") || " ");
        });
        
        if (rowData.length > 0) {
            tableData.push(rowData);
            maxCols = Math.max(maxCols, rowData.length);
        }
    });

    if (tableData.length === 0 || maxCols === 0) return "";

    // 确保所有行都有相同的列数
    tableData.forEach(row => {
        while (row.length < maxCols) {
            row.push(" ");
        }
    });

    // 生成Markdown表格
    let markdownTable = "";
    
    // 表格头部（第一行）
    if (tableData.length > 0) {
        markdownTable += "| " + tableData[0].join(" | ") + " |\n";
        
        // 分隔行
        markdownTable += "|" + " --- |".repeat(maxCols) + "\n";
        
        // 表格数据行（从第二行开始）
        for (let i = 1; i < tableData.length; i++) {
            markdownTable += "| " + tableData[i].join(" | ") + " |\n";
        }
    }
    
    return markdownTable;
}

/** Parse styles.xml to create a style mapping
 * @param {string} stylesXml The styles.xml content
 * @returns {Object} Style mapping object
 */
function parseStyles(stylesXml) {
    const styleMap = {};
    const doc = parseString(stylesXml);
    const styles = doc.getElementsByTagName("w:style");
    
    Array.from(styles).forEach(style => {
        const styleId = style.getAttribute("w:styleId");
        const styleName = style.getElementsByTagName("w:name")[0]?.getAttribute("w:val") || "";
        const styleType = style.getAttribute("w:type");
        
        if (styleId) {
            styleMap[styleId] = {
                name: styleName,
                type: styleType,
                isHeading: /^heading\s*\d+$/i.test(styleName) || /^title$/i.test(styleName)
            };
            
            // Extract heading level
            const headingMatch = styleName.match(/heading\s*(\d+)/i);
            if (headingMatch) {
                styleMap[styleId].headingLevel = parseInt(headingMatch[1]);
            } else if (/^title$/i.test(styleName)) {
                styleMap[styleId].headingLevel = 1;
            }
        }
    });
    
    return styleMap;
}

/** Get paragraph style information
 * @param {Element} paragraphNode The w:p element
 * @param {Object} styleMap Style mapping object
 * @returns {Object} Style information
 */
function getParagraphStyle(paragraphNode, styleMap) {
    const pPr = paragraphNode.getElementsByTagName("w:pPr")[0];
    if (!pPr) return {};
    
    const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
    if (!pStyle) return {};
    
    const styleId = pStyle.getAttribute("w:val");
    return styleMap[styleId] || {};
}

/** Get text formatting information from run properties
 * @param  runNode The w:r element
 * @returns {Object} Formatting information
 */
function getTextFormatting(runNode) {
    const rPr = runNode.getElementsByTagName("w:rPr")[0];
    if (!rPr) return {};
    
    return {
        bold: !!rPr.getElementsByTagName("w:b")[0],
        italic: !!rPr.getElementsByTagName("w:i")[0],
        underline: !!rPr.getElementsByTagName("w:u")[0],
        strike: !!rPr.getElementsByTagName("w:strike")[0]
    };
}

/** Apply Markdown formatting to text based on run properties
 * @param {string} text The text content
 * @param {Object} formatting Formatting information
 * @returns {string} Markdown formatted text
 */
function applyMarkdownFormatting(text, formatting) {
    let result = text;
    
    if (formatting.bold && formatting.italic) {
        result = `***${result}***`;
    } else if (formatting.bold) {
        result = `**${result}**`;
    } else if (formatting.italic) {
        result = `*${result}*`;
    }
    
    if (formatting.strike) {
        result = `~~${result}~~`;
    }
    
    // Note: Markdown doesn't have native underline support
    // You could use HTML tags if needed: <u>text</u>
    
    return result;
}

/** Apply paragraph-level formatting (headers)
 * @param {string} text The paragraph text
 * @param {Object} style Style information
 * @returns {string} Formatted paragraph
 */
function applyParagraphFormatting(text, style) {
    if (style.isHeading && style.headingLevel) {
        const headerPrefix = '#'.repeat(Math.min(style.headingLevel, 6));
        return `${headerPrefix} ${text}`;
    }
    
    return text;
}

/** Main function for parsing text from PowerPoint files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parsePowerPoint(file, callback, config) {
    // Files regex that hold our content of interest
    const allFilesRegex = /ppt\/(notesSlides|slides)\/(notesSlide|slide)\d+.xml/g;
    const slidesRegex   = /ppt\/slides\/slide\d+.xml/g;
    const slideNumberRegex = /lide(\d+)\.xml/;

    extractFiles(file, x => !!x.match(config.ignoreNotes ? slidesRegex : allFilesRegex))
        .then(files => {
            // Sort files by slide number and their notes (if any).
            files.sort((a, b) => {
                const matchedANumber = parseInt(a.path.match(slideNumberRegex)?.at(1), 10);
                const matchedBNumber = parseInt(b.path.match(slideNumberRegex)?.at(1), 10);

                const aNumber = isNaN(matchedANumber) ? Infinity : matchedANumber;
                const bNumber = isNaN(matchedBNumber) ? Infinity : matchedBNumber;

                return aNumber - bNumber || Number(a.path.includes('notes')) - Number(b.path.includes('notes'));
            });

            // Verify if atleast the slides xml files exist in the extracted files list.
            if (files.length == 0 || !files.map(file => file.path).some(filename => filename.match(slidesRegex)))
                throw ERRORMSG.fileCorrupted(file);

            // Check if any sorting is required.
            if (!config.ignoreNotes && config.putNotesAtLast)
                // Sort files according to previous order of taking text out of ppt/slides followed by ppt/notesSlides
                // For this we are looking at the index of notes which results in -1 in the main slide file and exists at a certain index in notes file names.
                files.sort((a, b) => a.path.indexOf("notes") - b.path.indexOf("notes"));

            // Returning an array of all the xml contents read using fs.readFileSync
            return files.map(file => ({ content: file.content, path: file.path }));
        })
        // ******************************** powerpoint xml files explanation ************************************
        // Structure of xmlContent of a powerpoint file is simple.
        // There are multiple xml files for each slide and correspondingly their notesSlide files.
        // All text nodes are within a:t tags and each of the text nodes that belong in one paragraph are clubbed together within a a:p tag.
        // So, we will filter out all the empty a:p tags and then combine all the a:t tag text inside for creating our response text.
        // ******************************************************************************************************
        .then(xmlContentArray => {
            /** Store all the markdown content to respond */
            let markdownContent = [];
            let currentSlideNumber = 0;
            let isProcessingNotes = false;

            xmlContentArray.forEach(xmlContentObj => {
                const { content: xmlContent, path } = xmlContentObj;
                
                // Extract slide number from path
                const slideMatch = path.match(slideNumberRegex);
                const slideNumber = slideMatch ? parseInt(slideMatch[1], 10) : 0;
                
                // Check if this is a notes slide
                const isNotesSlide = path.includes('notes');
                
                // If we're starting a new slide (not notes), add slide header
                if (!isNotesSlide && slideNumber !== currentSlideNumber) {
                    currentSlideNumber = slideNumber;
                    markdownContent.push(`\n## 幻灯片 ${slideNumber}\n`);
                    isProcessingNotes = false;
                }
                
                // If this is a notes slide and we haven't added notes header yet
                if (isNotesSlide && !isProcessingNotes) {
                    markdownContent.push(`\n### 备注\n`);
                    isProcessingNotes = true;
                }

                /** Find text nodes with a:p tags */
                const xmlParagraphNodesList = parseString(xmlContent).getElementsByTagName("a:p");
                
                /** Extract and format paragraph content */
                const paragraphContent = Array.from(xmlParagraphNodesList)
                    // Filter paragraph nodes than do not have any text nodes which are identifiable by a:t tag
                    .filter(paragraphNode => paragraphNode.getElementsByTagName("a:t").length != 0)
                    .map(paragraphNode => {
                        /** Find text nodes with a:t tags */
                        const xmlTextNodeList = paragraphNode.getElementsByTagName("a:t");
                        const paragraphText = Array.from(xmlTextNodeList)
                                .filter(textNode => textNode.childNodes[0] && textNode.childNodes[0].nodeValue)
                                .map(textNode => textNode.childNodes[0].nodeValue)
                                .join("");
                        
                        // Format as markdown list item if it's slide content (not notes)
                        return isNotesSlide ? paragraphText : `- ${paragraphText}`;
                    })
                    .join(config.newlineDelimiter ?? "\n");
                
                if (paragraphContent.trim()) {
                    markdownContent.push(paragraphContent);
                }
            });

            // Join all markdown content and clean up extra newlines
            const finalMarkdown = markdownContent
                .join(config.newlineDelimiter ?? "\n")
                .replace(/\n{3,}/g, '\n\n') // Replace multiple newlines with double newlines
                .trim();

            // Respond by calling the Callback function with markdown content
            callback(finalMarkdown, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Main function for parsing text from Excel files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseExcel(file, callback, config) {
    // Files regex that hold our content of interest
    const sheetsRegex     = /xl\/worksheets\/sheet\d+.xml/g;
    const drawingsRegex   = /xl\/drawings\/drawing\d+.xml/g;
    const chartsRegex     = /xl\/charts\/chart\d+.xml/g;
    const stringsFilePath = 'xl/sharedStrings.xml';

    extractFiles(file, x => [sheetsRegex, drawingsRegex, chartsRegex].some(fileRegex => x.match(fileRegex)) || x == stringsFilePath)
        .then(files => {
            // Verify if atleast the slides xml files exist in the extracted files list.
            if (files.length == 0 || !files.map(file => file.path).some(filename => filename.match(sheetsRegex)))
                throw ERRORMSG.fileCorrupted(file);

            return {
                sheetFiles:        files.filter(file => file.path.match(sheetsRegex)).map((file, index) => ({content: file.content, index: index + 1})),
                drawingFiles:      files.filter(file => file.path.match(drawingsRegex)).map(file => file.content),
                chartFiles:        files.filter(file => file.path.match(chartsRegex)).map(file => file.content),
                sharedStringsFile: files.filter(file => file.path == stringsFilePath).map(file => file.content)[0],
            };
        })
        .then(xmlContentFilesObject => {
            /** Store all the markdown content to respond */
            let markdownContent = [];

            /** Function to check if the given c node is a valid inline string node. */
            function isValidInlineStringCNode(cNode) {
                // Initial check to see if the passed node is a cNode
                if (cNode.tagName.toLowerCase() != 'c')
                    return false;
                if (cNode.getAttribute("t") != 'inlineStr')
                    return false;
                const childNodesNamedIs = cNode.getElementsByTagName('is');
                if (childNodesNamedIs.length != 1)
                    return false;
                const childNodesNamedT = childNodesNamedIs[0].getElementsByTagName('t');
                if (childNodesNamedT.length != 1)
                    return false;
                return childNodesNamedT[0].childNodes[0] && childNodesNamedT[0].childNodes[0].nodeValue != '';
            }

            /** Function to check if the given c node has a valid v node */
            function hasValidVNodeInCNode(cNode) {
                return cNode.getElementsByTagName("v")[0]
                    && cNode.getElementsByTagName("v")[0].childNodes[0]
                    && cNode.getElementsByTagName("v")[0].childNodes[0].nodeValue != ''
            }

            /** Function to get cell reference (like A1, B2) from c node */
            function getCellReference(cNode) {
                return cNode.getAttribute('r') || '';
            }

            /** Function to convert cell reference to row and column numbers */
            function parseReference(ref) {
                const match = ref.match(/^([A-Z]+)(\d+)$/);
                if (!match) return { row: 0, col: 0 };
                
                const colStr = match[1];
                const rowNum = parseInt(match[2]);
                
                let colNum = 0;
                for (let i = 0; i < colStr.length; i++) {
                    colNum = colNum * 26 + (colStr.charCodeAt(i) - 64);
                }
                
                return { row: rowNum, col: colNum };
            }

            /** Find text nodes with t tags in sharedStrings xml file. If the sharedStringsFile is not present, we return an empty array. */
            const sharedStringsXmlTNodesList = xmlContentFilesObject.sharedStringsFile != undefined ? parseString(xmlContentFilesObject.sharedStringsFile).getElementsByTagName("t")
                                                                                                    : [];
            /** Create shared string array. This will be used as a map to get strings from within sheet files. */
            const sharedStrings = Array.from(sharedStringsXmlTNodesList)
                                    .map(tNode => tNode.childNodes[0]?.nodeValue ?? '');

            // Parse Sheet files and convert to markdown tables
            xmlContentFilesObject.sheetFiles.forEach(sheetData => {
                const sheetXmlContent = sheetData.content;
                const sheetIndex = sheetData.index;
                
                markdownContent.push(`## 工作表 ${sheetIndex}\n`);
                
                /** Find text nodes with c tags in sheet xml file */
                const sheetsXmlCNodesList = parseString(sheetXmlContent).getElementsByTagName("c");
                
                // Create a map to store cell data by position
                const cellData = new Map();
                let maxRow = 0;
                let maxCol = 0;
                
                // Process all cells and organize by position
                Array.from(sheetsXmlCNodesList)
                    .filter(cNode => isValidInlineStringCNode(cNode) || hasValidVNodeInCNode(cNode))
                    .forEach(cNode => {
                        const cellRef = getCellReference(cNode);
                        const { row, col } = parseReference(cellRef);
                        
                        if (row > 0 && col > 0) {
                            maxRow = Math.max(maxRow, row);
                            maxCol = Math.max(maxCol, col);
                            
                            let cellValue = '';
                            
                            // Processing if this is a valid inline string c node.
                            if (isValidInlineStringCNode(cNode)) {
                                cellValue = cNode.getElementsByTagName('is')[0].getElementsByTagName('t')[0].childNodes[0].nodeValue;
                            }
                            // Processing if this c node has a valid v node.
                            else if (hasValidVNodeInCNode(cNode)) {
                                /** Flag whether this node's value represents an index in the shared string array */
                                const isIndexInSharedStrings = cNode.getAttribute("t") == "s";
                                /** Find value nodes represented by v tags */
                                const value = parseInt(cNode.getElementsByTagName("v")[0].childNodes[0].nodeValue, 10);
                                // Validate text
                                if (isIndexInSharedStrings && value >= sharedStrings.length)
                                    throw ERRORMSG.fileCorrupted(file);

                                cellValue = isIndexInSharedStrings
                                        ? sharedStrings[value]
                                        : value.toString();
                            }
                            
                            cellData.set(`${row}-${col}`, cellValue);
                        }
                    });
                
                // Generate markdown table if we have data
                if (maxRow > 0 && maxCol > 0) {
                    // Create table header
                    let tableHeader = '|';
                    let tableSeparator = '|';
                    for (let col = 1; col <= maxCol; col++) {
                        const colLetter = String.fromCharCode(64 + col);
                        tableHeader += ` ${colLetter} |`;
                        tableSeparator += ' --- |';
                    }
                    
                    markdownContent.push(tableHeader);
                    markdownContent.push(tableSeparator);
                    
                    // Create table rows
                    for (let row = 1; row <= maxRow; row++) {
                        let tableRow = '|';
                        for (let col = 1; col <= maxCol; col++) {
                            const cellValue = cellData.get(`${row}-${col}`) || '';
                            // Escape markdown special characters in cell content
                            const escapedValue = cellValue.toString().replace(/\|/g, '\\|').replace(/\n/g, '<br>');
                            tableRow += ` ${escapedValue} |`;
                        }
                        markdownContent.push(tableRow);
                    }
                } else {
                    markdownContent.push('*此工作表为空*');
                }
                
                markdownContent.push(''); // Add empty line after each sheet
            });

            // Parse Drawing files
            if (xmlContentFilesObject.drawingFiles.length > 0) {
                markdownContent.push('## 绘图内容\n');
                
                xmlContentFilesObject.drawingFiles.forEach((drawingXmlContent, index) => {
                    /** Find text nodes with a:p tags */
                    const drawingsXmlParagraphNodesList = parseString(drawingXmlContent).getElementsByTagName("a:p");
                    
                    const drawingTexts = Array.from(drawingsXmlParagraphNodesList)
                        .filter(paragraphNode => paragraphNode.getElementsByTagName("a:t").length != 0)
                        .map(paragraphNode => {
                            /** Find text nodes with a:t tags */
                            const xmlTextNodeList = paragraphNode.getElementsByTagName("a:t");
                            return Array.from(xmlTextNodeList)
                                    .filter(textNode => textNode.childNodes[0] && textNode.childNodes[0].nodeValue)
                                    .map(textNode => textNode.childNodes[0].nodeValue)
                                    .join("");
                        })
                        .filter(text => text.trim() !== '');
                    
                    if (drawingTexts.length > 0) {
                        markdownContent.push(`### 绘图 ${index + 1}\n`);
                        drawingTexts.forEach(text => {
                            markdownContent.push(`- ${text}`);
                        });
                        markdownContent.push(''); // Add empty line
                    }
                });
            }

            // Parse Chart files
            if (xmlContentFilesObject.chartFiles.length > 0) {
                markdownContent.push('## 图表数据\n');
                
                xmlContentFilesObject.chartFiles.forEach((chartXmlContent, index) => {
                    /** Find text nodes with c:v tags */
                    const chartsXmlCVNodesList = parseString(chartXmlContent).getElementsByTagName("c:v");
                    
                    const chartValues = Array.from(chartsXmlCVNodesList)
                        .filter(cVNode => cVNode.childNodes[0] && cVNode.childNodes[0].nodeValue)
                        .map(cVNode => cVNode.childNodes[0].nodeValue)
                        .filter(value => value.trim() !== '');
                    
                    if (chartValues.length > 0) {
                        markdownContent.push(`### 图表 ${index + 1}\n`);
                        chartValues.forEach(value => {
                            markdownContent.push(`- ${value}`);
                        });
                        markdownContent.push(''); // Add empty line
                    }
                });
            }

            // Respond by calling the Callback function with markdown content
            const finalMarkdown = markdownContent.join(config.newlineDelimiter ?? "\n");
            callback(finalMarkdown, undefined);
        })
        .catch(e => callback(undefined, e));
}


/** Main function for parsing text from open office files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {void}
 */
function parseOpenOffice(file, callback, config) {
    /** The target content xml file for the openoffice file. */
    const mainContentFilePath     = 'content.xml';
    const objectContentFilesRegex = /Object \d+\/content.xml/g;

    extractFiles(file, x => x == mainContentFilePath || !!x.match(objectContentFilesRegex))
        .then(files => {
            // Verify if atleast the content xml file exists in the extracted files list.
            if (!files.map(file => file.path).includes(mainContentFilePath))
                throw ERRORMSG.fileCorrupted(file);

            return {
                mainContentFile:    files.filter(file => file.path == mainContentFilePath).map(file => file.content)[0],
                objectContentFiles: files.filter(file => file.path.match(objectContentFilesRegex)).map(file => file.content),
            }
        })
        // ********************************** openoffice xml files explanation **********************************
        // Structure of xmlContent of openoffice files is simple.
        // All text nodes are within text:h and text:p tags with all kinds of formatting within nested tags.
        // All text in these tags are separated by new line delimiters.
        // Objects like charts in ods files are in Object d+/content.xml with the same way as above.
        // ******************************************************************************************************
        .then(xmlContentFilesObject => {
            /** Store all the notes text content to respond */
            let notesText = [];
            /** Store all the text content to respond */
            let responseText = [];

            /** List of allowed text tags */
            const allowedTextTags = ["text:p", "text:h"];
            /** List of notes tags */
            const notesTag = "presentation:notes";

            /** Parse OpenOffice table and convert to Markdown table */
            function parseOpenOfficeTable(tableElement) {
                const rows = tableElement.getElementsByTagName("table:table-row");
                if (rows.length === 0) return "";

                const tableData = [];
                let maxCols = 0;

                // Extract data from each row
                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    const cells = row.getElementsByTagName("table:table-cell");
                    const rowData = [];

                    for (let j = 0; j < cells.length; j++) {
                        const cell = cells[j];
                        const cellContent = [];
                        
                        // Extract text from paragraphs within the cell
                        const paragraphs = cell.getElementsByTagName("text:p");
                        for (let k = 0; k < paragraphs.length; k++) {
                            const paragraphText = extractTextFromNode(paragraphs[k]);
                            if (paragraphText.trim()) {
                                cellContent.push(paragraphText.trim());
                            }
                        }
                        
                        // Handle repeated columns (table:number-columns-repeated)
                        const colsRepeated = cell.getAttribute("table:number-columns-repeated");
                        const repeatCount = colsRepeated ? parseInt(colsRepeated, 10) : 1;
                        
                        const cellText = cellContent.join(" ") || " ";
                        for (let r = 0; r < repeatCount; r++) {
                            rowData.push(cellText);
                        }
                    }
                    
                    if (rowData.length > 0) {
                        tableData.push(rowData);
                        maxCols = Math.max(maxCols, rowData.length);
                    }
                }

                if (tableData.length === 0) return "";

                // Normalize all rows to have the same number of columns
                tableData.forEach(row => {
                    while (row.length < maxCols) {
                        row.push(" ");
                    }
                });

                // Generate markdown table
                let markdownTable = "\n";
                
                // Header row
                markdownTable += "|" + tableData[0].map(cell => {
                    // Escape markdown special characters
                    const escapedCell = cell.replace(/\|/g, '\\|').replace(/\n/g, '<br>');
                    return ` ${escapedCell} `;
                }).join("|") + "|\n";
                
                // Separator row
                markdownTable += "|" + Array(maxCols).fill(" --- ").join("|") + "|\n";
                
                // Data rows
                for (let i = 1; i < tableData.length; i++) {
                    markdownTable += "|" + tableData[i].map(cell => {
                        // Escape markdown special characters
                        const escapedCell = cell.replace(/\|/g, '\\|').replace(/\n/g, '<br>');
                        return ` ${escapedCell} `;
                    }).join("|") + "|\n";
                }
                
                markdownTable += "\n";
                return markdownTable;
            }

            /** Extract text content from a node recursively */
            function extractTextFromNode(node) {
                let text = "";
                if (node.nodeType === 3) { // Text node
                    text += node.nodeValue || "";
                } else if (node.nodeType === 1) { // Element node
                    for (let i = 0; i < node.childNodes.length; i++) {
                        text += extractTextFromNode(node.childNodes[i]);
                    }
                }
                return text;
            }

            /** Main dfs traversal function that goes from one node to its children and returns the value out. */
            function extractAllTextsFromNode(root) {
                let xmlTextArray = []
                for (let i = 0; i < root.childNodes.length; i++)
                    traversal(root.childNodes[i], xmlTextArray, true, root.tagName);
                return xmlTextArray.join("");
            }
            
            /** Traversal function that gets recursive calling. */
            function traversal(node, xmlTextArray, isFirstRecursion, parentTagName) {
                if (!node.childNodes || node.childNodes.length == 0) {
                    if (node.parentNode.tagName.indexOf('text') == 0 && node.nodeValue) {
                        if (isNotesNode(node.parentNode) && (config.putNotesAtLast || config.ignoreNotes)) {
                            notesText.push(node.nodeValue);
                            if (allowedTextTags.includes(node.parentNode.tagName) && !isFirstRecursion)
                                notesText.push(config.newlineDelimiter ?? "\n");
                        }
                        else {
                            xmlTextArray.push(node.nodeValue);
                            if (allowedTextTags.includes(node.parentNode.tagName) && !isFirstRecursion)
                                xmlTextArray.push(config.newlineDelimiter ?? "\n");
                        }
                    }
                    return;
                }

                for (let i = 0; i < node.childNodes.length; i++)
                    traversal(node.childNodes[i], xmlTextArray, false, parentTagName);
            }

            /** Checks if the given node has an ancestor which is a notes tag. We use this information to put the notes in the response text and its position. */
            function isNotesNode(node) {
                if (node.tagName == notesTag)
                    return true;
                if (node.parentNode)
                    return isNotesNode(node.parentNode);
                return false;
            }

            /** Checks if the given node has an ancestor which is also an allowed text tag. In that case, we ignore the child text tag. */
            function isInvalidTextNode(node) {
                if (allowedTextTags.includes(node.tagName))
                    return true;
                if (node.parentNode)
                    return isInvalidTextNode(node.parentNode);
                return false;
            }

            /** Convert text to markdown format based on tag type */
            function formatAsMarkdown(text, tagName) {
                if (!text.trim()) return '';
                
                if (tagName === 'text:h') {
                    // Convert headings to markdown format
                    return `## ${text.trim()}\n\n`;
                } else if (tagName === 'text:p') {
                    // Convert paragraphs to markdown format
                    return `${text.trim()}\n\n`;
                }
                return text;
            }

            /** The xml string parsed as xml array */
            const xmlContentArray = [xmlContentFilesObject.mainContentFile, ...xmlContentFilesObject.objectContentFiles].map(xmlContent => parseString(xmlContent));
            
            // Iterate over each xmlContent and extract text from them.
            xmlContentArray.forEach(xmlContent => {
                // First, process tables
                const tables = xmlContent.getElementsByTagName("table:table");
                for (let i = 0; i < tables.length; i++) {
                    const markdownTable = parseOpenOfficeTable(tables[i]);
                    if (markdownTable.trim()) {
                        responseText.push(markdownTable);
                    }
                }

                // Then process regular text nodes
                /** Find text nodes with text:h and text:p tags in xmlContent */
                const xmlTextNodesList = [...Array.from(xmlContent
                                                .getElementsByTagName("*"))
                                                .filter(node => allowedTextTags.includes(node.tagName)
                                                    && !isInvalidTextNode(node.parentNode)
                                                    && !isInsideTable(node))
                                            ];
                
                /** Check if a node is inside a table */
                function isInsideTable(node) {
                    let parent = node.parentNode;
                    while (parent) {
                        if (parent.tagName === "table:table") {
                            return true;
                        }
                        parent = parent.parentNode;
                    }
                    return false;
                }
                
                const markdownContent = xmlTextNodesList
                    // Add every text information from within this textNode and combine them together.
                    .map(textNode => {
                        const text = extractAllTextsFromNode(textNode);
                        return formatAsMarkdown(text, textNode.tagName);
                    })
                    .filter(text => text.trim() !== "")
                    .join('');
                    
                if (markdownContent.trim()) {
                    responseText.push(markdownContent);
                }
            });

            // Add notes text at the end if the user config says so.
            if (!config.ignoreNotes && config.putNotesAtLast) {
                const notesMarkdown = notesText.join('').trim();
                if (notesMarkdown) {
                    responseText.push(`\n## 备注\n\n${notesMarkdown}\n`);
                }
            }

            // Respond by calling the Callback function.
            const finalMarkdown = responseText.join('').trim();
            callback(finalMarkdown, undefined);
        })
        .catch(e => callback(undefined, e));
}



/** Helper function to convert text to markdown format
 * @param {string} text - The raw text content
 * @returns {string} - Formatted markdown text
 */
function convertToMarkdown(text) {
    // Split text into lines
    const lines = text.split('\n');
    let markdownText = '';
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Skip empty lines
        if (!line) {
            markdownText += '\n';
            continue;
        }
        
        let isHeading = false;
        
        // Method 1: All caps lines (更宽松)
        if (line.match(/^[A-Z\s\d\.\-\(\)\:]+$/) && 
            line.length > 2 && 
            line.length < 100 && 
            !line.includes('://')) { // 只排除URL
            
            const nextLine = i + 1 < lines.length ? lines[i + 1].trim() : '';
            const prevLine = i > 0 ? lines[i - 1].trim() : '';
            
            // 更宽松的上下文检查
            if ((!prevLine || prevLine.length === 0) || 
                (!nextLine || nextLine.length === 0) ||
                i === 0 || i === lines.length - 1 ||
                (nextLine.length === 0) || // 只要后面有空行就可以
                (prevLine.length > 0 && nextLine.length > 0 && line.length < 30)) { // 短行也可以
                isHeading = true;
            }
        }
        
        // Method 2: 数字编号模式（更宽松）
        if (line.match(/^\d+(\.\d+)*[\.\)]?\s*[A-Za-z\u4e00-\u9fff]/) && 
            line.length < 120) {
            isHeading = true;
        }
        
        // Method 3: 常见标题词（英文+中文）
        if (line.match(/^(Chapter|Section|Part|Appendix|Introduction|Conclusion|Summary|Overview|Abstract|Background|Method|Result|Discussion|Example|Sample|Demo|Tutorial|Guide|Manual|Reference|API|Usage|Installation|Configuration|Setup|第.*章|第.*节|第.*部分|章节|摘要|总结|概述|介绍|结论|背景|方法|结果|讨论|示例|样例|演示|教程|指南|手册|参考|用法|安装|配置|设置|前言|序言|目录|附录|说明|描述|定义|原理|实现|应用|测试|验证|分析|评估|比较|优化|改进|扩展|未来|展望|致谢|参考文献)\s+/i) &&
            line.length < 100) {
            isHeading = true;
        }
        
        // Method 4: 标题格式（更宽松，支持中文）
        if (line.match(/^[A-Z\u4e00-\u9fff][A-Za-z\u4e00-\u9fff\s\d\-\.\:]+$/) && 
            line.length > 3 && 
            line.length < 80 &&
            !line.endsWith('.') &&
            !line.includes('://') &&
            line.split(/\s+/).length <= 12) { // 允许更多单词
            
            const nextLine = i + 1 < lines.length ? lines[i + 1].trim() : '';
            const prevLine = i > 0 ? lines[i - 1].trim() : '';
            
            // 更宽松的条件
            if ((!prevLine || prevLine.length === 0) || 
                (!nextLine || nextLine.length === 0) ||
                i === 0 || i === lines.length - 1 ||
                (line.length < 40)) { // 短行更容易被识别为标题
                isHeading = true;
            }
        }
        
        // Method 5: 混合大小写标题（新增，支持中文）
        if (line.match(/^[A-Z\u4e00-\u9fff][a-z\u4e00-\u9fff]+(?:\s+[A-Za-z\u4e00-\u9fff][a-z\u4e00-\u9fff]*)*$/) && 
            line.length > 4 && 
            line.length < 70 &&
            !line.endsWith('.') &&
            !line.match(/^(The|A|An|This|That|These|Those|这|那|这些|那些|本|该)\s/)) {
            
            const nextLine = i + 1 < lines.length ? lines[i + 1].trim() : '';
            
            // 如果后面有空行或文档结束
            if (!nextLine || nextLine.length === 0 || i === lines.length - 1) {
                isHeading = true;
            }
        }
        
        // Method 6: 短行标题检测（新增，更激进，支持中文）
        if (line.length > 2 && line.length < 25 && 
            line.match(/^[A-Za-z\u4e00-\u9fff][A-Za-z\u4e00-\u9fff\s\d\-]*$/) &&
            !line.endsWith('.') &&
            !line.includes('://')) {
            
            const nextLine = i + 1 < lines.length ? lines[i + 1].trim() : '';
            const prevLine = i > 0 ? lines[i - 1].trim() : '';
            
            // 如果是独立的短行
            if ((!nextLine || nextLine.length === 0) && 
                (!prevLine || prevLine.length === 0 || prevLine.length > line.length * 2)) {
                isHeading = true;
            }
        }
        
        // Method 7: 冒号结尾的标题（新增，支持中文）
        if (line.endsWith(':') && 
            line.length > 3 && 
            line.length < 60 &&
            line.match(/^[A-Za-z\u4e00-\u9fff][A-Za-z\u4e00-\u9fff\s\d\-]*:$/)) {
            isHeading = true;
        }
        
        // Method 8: 中文特有标题模式（新增）
        if (line.match(/^[一二三四五六七八九十百千万]+[、．\.]\s*[\u4e00-\u9fff]/) && 
            line.length < 80) {
            isHeading = true;
        }
        
        // Method 9: 中文括号编号（新增）
        if (line.match(/^[（\(][一二三四五六七八九十\d]+[）\)]\s*[\u4e00-\u9fff]/) && 
            line.length < 80) {
            isHeading = true;
        }
        
        if (isHeading) {
            markdownText += `## ${line}\n\n`;
            continue;
        }
        
        // Detect numbered lists
        if (line.match(/^\d+[\.\)]\s/)) {
            markdownText += `${line}\n`;
            continue;
        }
        
        // Detect bullet points
        if (line.match(/^[\-\*\•]\s/)) {
            markdownText += `- ${line.substring(2)}\n`;
            continue;
        }
        
        // Regular paragraph text
        markdownText += `${line}\n\n`;
    }
    
    return markdownText.trim();
}

/** Main function for parsing text from pdf files
 * @param {string | Buffer}    file     File path or Buffers
 * @param {function}           callback Callback function that returns value or error
 * @param {OfficeParserConfig} config   Config Object for officeParser
 * @returns {Promise<void>}
 */
async function parsePdf(file, callback, config) {
    // Wait for pdfjs module to be loaded once
    const pdfjs = await pdfjsPromise;

    // Get the pdfjs document for the filepath or Uint8Array buffers.
    // pdfjs does not accept Buffers directly, so we convert them to Uint8Array.
    pdfjs.getDocument(file instanceof Buffer ? new Uint8Array(file) : file).promise
        // We go through each page and build our text content promise array.
        .then(document => Promise.all(Array.from({ length: document.numPages }, (_, index) => document.getPage(index + 1).then(page => page.getTextContent()))))
        // Each textContent item has property 'items' which is an array of objects.
        // Each object element in the array has text stored in their 'str' key.
        // The concatenation of str is what makes our pdf content.
        // str already contains any space that was in the text.
        // So, we only care about when to add the new line.
        // That we determine using transform[5] value which is the y-coordinate of the item object.
        // So, if there is a mismatch in the transform[5] value between the current item and the previous item, we put a line break.
        .then(textContentArray => {
            /** Store all the text content to respond */
            const responseText = textContentArray
                                    .map(textContent => textContent.items)      // Get all the items
                                    .flat()                                     // Flatten all the items object
                                    .reduce((a, v) =>  (
                                        // the items could be TextItem or a TextMarkedContent.
                                        // We are only interested in the TextItem which has a str property.
                                        'str' in v && v.str != ''
                                            ? {
                                                text: a.text + (v.transform[5] != a.transform5 ? (config.newlineDelimiter ?? "\n") : '') + v.str,
                                                transform5: v.transform[5]
                                            } : {
                                                text: a.text,
                                                transform5: a.transform5
                                            }
                                    ),
                                    {
                                        text: '',
                                        transform5: undefined
                                    }).text;

            // Convert the extracted text to markdown format
            const markdownText = convertToMarkdown(responseText);
            
            callback(markdownText, undefined);
        })
        .catch(e => callback(undefined, e));
}

/** Main async function with callback to execute parseOffice for supported files
 * @param {string | Buffer | ArrayBuffer} srcFile      File path or file buffers or Javascript ArrayBuffer
 * @param {function}                      callback     Callback function that returns value or error
 * @param {OfficeParserConfig}            [config={}]  [OPTIONAL]: Config Object for officeParser
 * @returns {void}
 */
function parseOffice(srcFile, callback, config = {}) {
    // Make a clone of the config with default values such that none of the config flags are undefined.
    /** @type {OfficeParserConfig} */
    const internalConfig = {
        ignoreNotes: false,
        newlineDelimiter: '\n',
        putNotesAtLast: false,
        outputErrorToConsole: false,
        ...config
    };

    // Our internal code can process regular node Buffers or file path.
    // So, if the src file was presented as ArrayBuffers, we create Buffers from them.
    let file = srcFile instanceof ArrayBuffer ? Buffer.from(srcFile)
                                              : srcFile;

    /**
     * Prepare file for processing
     * @type {Promise<{ file:string | Buffer, ext: string}>}
     */
    const filePreparedPromise = new Promise((res, rej) => {
        // Check if buffer
        if (Buffer.isBuffer(file))
            // Guess file type from buffer
            return fileType.fromBuffer(file)
                .then(data => res({ file: file, ext: data.ext.toLowerCase() }))
                .catch(() => rej(ERRORMSG.improperBuffers));
        else if (typeof file === 'string') {
            // Not buffers but real file path.
            // Check if file exists
            if (!fs.existsSync(file))
                throw ERRORMSG.fileDoesNotExist(file);

            // resolve promise
            res({ file: file, ext: file.split(".").pop() });
        }
        else
            rej(ERRORMSG.invalidInput);
    });

    // Process filePreparedPromise resolution.
    filePreparedPromise
        .then(({ file, ext }) => {
            // Switch between parsing functions depending on extension.
            switch (ext) {
                case "docx":
                    parseWord(file, internalCallback, internalConfig);
                    break;
                case "pptx":
                    parsePowerPoint(file, internalCallback, internalConfig);
                    break;
                case "xlsx":
                    parseExcel(file, internalCallback, internalConfig);
                    break;
                case "odt":
                case "odp":
                case "ods":
                    parseOpenOffice(file, internalCallback, internalConfig);
                    break;
                case "pdf":
                    parsePdf(file, internalCallback, internalConfig);
                    break;

                default:
                    internalCallback(undefined, ERRORMSG.extensionUnsupported(ext));  // Call the internalCallback function which removes the temp files if required.
            }

            /** Internal callback function that calls the user's callback function passed in argument and removes the temp files if required */
            function internalCallback(data, err) {
                // Check if there is an error. Throw if there is an error.
                if (err)
                    return handleError(err, callback, internalConfig.outputErrorToConsole);

                // Call the original callback
                callback(data, undefined);
            }
        })
        .catch(error => handleError(error, callback, internalConfig.outputErrorToConsole));
}

/** Main async function that can be used with await to execute parseOffice. Or it can be used with promises.
 * @param {string | Buffer | ArrayBuffer} srcFile     File path or file buffers or Javascript ArrayBuffer
 * @param {OfficeParserConfig}            [config={}] [OPTIONAL]: Config Object for officeParser
 * @returns {Promise<string>}
 */
function parseOfficeAsync(srcFile, config = {}) {
    return new Promise((res, rej) => {
        parseOffice(srcFile, function (data, err) {
            if (err)
                return rej(err);
            return res(data);
        }, config);
    });
}

/** Extract specific files from either a ZIP file buffer or file path based on a filter function.
 * @param {Buffer|string}          zipInput ZIP file input, either a Buffer or a file path (string).
 * @param {(x: string) => boolean} filterFn A function that receives the entry object and returns true if the file should be extracted.
 * @returns {Promise<{ path: string, content: string }[]>} Resolves to an array of object 
 */
function extractFiles(zipInput, filterFn) {
    return new Promise((res, rej) => {
        /** Processes zip file and resolves with the path of file and their content.
         * @param {yauzl.ZipFile} zipfile
         */
        const processZipfile = (zipfile) => {
            /** @type {{ path: string, content: string }[]} */
            const extractedFiles = [];
            zipfile.readEntry();

            /** @param {yauzl.Entry} entry  */
            function processEntry(entry) {
                // Use the filter function to determine if the file should be extracted
                if (filterFn(entry.fileName)) {
                    zipfile.openReadStream(entry, (err, readStream) => {
                        if (err)
                            return rej(err);

                        // Use concat-stream to collect the data into a single Buffer
                        readStream.pipe(concat(data => {
                            extractedFiles.push({
                                path: entry.fileName,
                                content: data.toString()
                            });
                            zipfile.readEntry(); // Continue reading entries
                        }));
                    });
                }
                else
                    zipfile.readEntry(); // Skip entries that don't match the filter
            }

            zipfile.on('entry', processEntry);
            zipfile.on('end', () => res(extractedFiles));
            zipfile.on('error', rej);
        };

        // Determine whether the input is a buffer or file path
        if (Buffer.isBuffer(zipInput)) {
            // Process ZIP from Buffer
            yauzl.fromBuffer(zipInput, { lazyEntries: true }, (err, zipfile) => {
                if (err) return rej(err);
                processZipfile(zipfile);
            });
        }
        else if (typeof zipInput === 'string') {
            // Process ZIP from File Path
            yauzl.open(zipInput, { lazyEntries: true }, (err, zipfile) => {
                if (err) return rej(err);
                processZipfile(zipfile);
            });
        }
        else
            rej(ERRORMSG.invalidInput);
    });
}

/** Handle error by logging it to console if permitted by the config.
 * And after that, trigger the callback function with the error value.
 * @param {string}   error                Error text
 * @param {function} callback             Callback function provided by the caller
 * @param {boolean}  outputErrorToConsole Flag to log error to console.
 * @returns {void}
 */
function handleError(error, callback, outputErrorToConsole) {
    if (error && outputErrorToConsole)
        console.error(ERRORHEADER + error);

    callback(undefined, new Error(ERRORHEADER + error));
}


// Export functions
module.exports.parseOffice      = parseOffice;
module.exports.parseOfficeAsync = parseOfficeAsync;


// Run this library on CLI
if ((typeof process.argv[0] == 'string' && (process.argv[0].split('/').pop() == "node" || process.argv[0].split('/').pop() == "npx")) &&
    (typeof process.argv[1] == 'string' && (process.argv[1].split('/').pop() == "officeParser.js" || process.argv[1].split('/').pop().toLowerCase() == "officeparser"))) {

    // Extract arguments after the script is called
    /** Stores the list of arguments for this CLI call
     * @type {string[]}
     */
    const args = process.argv.slice(2);
    /** Stores the file argument for this CLI call
     * @type {string | Buffer | undefined}
     */
    let fileArg = undefined;
    /** Stores the config arguments for this CLI call
     * @type {string[]}
     */
    const configArgs = [];

    /** Function to identify if an argument is a config option (i.e., --key=value)
     * @param {string} arg Argument passed in the CLI call.
     */ 
    function isConfigOption(arg) {
        return arg.startsWith('--') && arg.includes('=');
    }

    // Loop through arguments to separate file path and config options
    args.forEach(arg => {
        if (isConfigOption(arg))
            // It's a config option
            configArgs.push(arg);
        else if (!fileArg)
            // First non-config argument is assumed to be the file path
            fileArg = arg;
    });

    // Check if we have a valid file argument
    // If not, we return error and we write the instructions on how to use the library on the terminal.
    if (fileArg != undefined) {
        /** Helper function to parse config arguments from CLI
         * @param {string[]} args List of string arguments that we need to parse to understand the config flag they represent.
         */
        function parseCLIConfigArgs(args) {
            /** @type {OfficeParserConfig} */
            const config = {};
            args.forEach(arg => {
                // Split the argument by '=' to differentiate between the key and value
                const [key, value] = arg.split('=');

                // We only care about the keys that are important to us. We ignore any other key.
                switch (key) {
                    case '--ignoreNotes':
                        config.ignoreNotes = value.toLowerCase() === 'true';
                        break;
                    case '--newlineDelimiter':
                        config.newlineDelimiter = value;
                        break;
                    case '--putNotesAtLast':
                        config.putNotesAtLast = value.toLowerCase() === 'true';
                        break;
                    case '--outputErrorToConsole':
                        config.outputErrorToConsole = value.toLowerCase() === 'true';
                        break;
                }
            });

            return config;
        }

        // Parse CLI config arguments
        const config = parseCLIConfigArgs(configArgs);

        // Execute parseOfficeAsync with file and config
        parseOfficeAsync(fileArg, config)
            .then(text => console.log(text))
            .catch(error => console.error(ERRORHEADER + error));
    }
    else {
        console.error(ERRORMSG.improperArguments);

        const CLI_INSTRUCTIONS =
`
=== How to Use officeParser CLI ===

Usage:
    node officeparser [--configOption=value] [FILE_PATH]

Example:
    node officeparser --ignoreNotes=true --putNotesAtLast=true ./example.docx

Config Options:
    --ignoreNotes=[true|false]          Flag to ignore notes from files like PowerPoint. Default is false.
    --newlineDelimiter=[delimiter]      The delimiter to use for new lines. Default is '\\n'.
    --putNotesAtLast=[true|false]       Flag to collect notes at the end of files like PowerPoint. Default is false.
    --outputErrorToConsole=[true|false] Flag to output errors to the console. Default is false.

Note:
    The order of file path and config options doesn't matter.
`;
        // Usage instructions for the user
        console.log(CLI_INSTRUCTIONS);
    }
}