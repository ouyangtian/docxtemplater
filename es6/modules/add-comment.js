const wrapper = require("../module-wrapper");
const traits = require("../traits");
const { DOMParser, XMLSerializer } = require("xmldom");
const { concatArrays, isContent, getLeft, getRight, getNearestLeft, getNearestRight } = require("../doc-utils");
const { throwRawTagShouldBeOnlyTextInParagraph, throwMalformedXml } = require("../errors");

const re = /\<comment\s+value=['"]([\w\t]+)['"]\>([^<]+)\<\/comment\>/g

class AddCommentModule {
    optionsTransformer(options, docxtemplater) {
        this.data = docxtemplater.data;
        this.docxtemplater = docxtemplater;
        this.commentCount = 0
        this.comments = {}

        let modules = docxtemplater.modules
        for (let i=0; i<modules.length; ++i) {
            let module = modules[i]
            if (module.name && module.name == "RawXmlModule") {
                this.rawxml = module
                break
            }
        };
        return options;
    }

    // 将含有comment标记的数据转换成rawxml
    convertToRawXml(memberData) {
        // 由于当前字段需要标注，因此需要将所有对应的值都转换为rawxml类型
        let xmlObjectArray = null
        let segment = []

        if (xmlObjectArray = this.getXmlObject(memberData)) {
            // 根据xmlObject信息转换成rawxml
            let pos = 0;

            for (let j=0; j<xmlObjectArray.length; ++j) {
                let xmlObject = xmlObjectArray[j];

                if (pos != xmlObject.start) {
                    segment.push({
                        start: pos,
                        end: xmlObject.start,
                    });
                } 

                segment.push({
                    start: xmlObject.start,
                    end: xmlObject.end,
                    xml: xmlObject
                });

                pos = xmlObject.end
            }

            // 处理最后剩下的部分
            let last = segment[segment.length-1] 
            if (last.end != memberData.length) {
                segment.push({
                    start: last.end,
                    end: memberData.length,
                });
            }
        } else {
            segment.push({
                start: 0,
                end: memberData.length
            });
        }

        // 将当前data数据转换成rawxml
        let _this = this
        let result = segment.reduce(function(str, part) {
            if (!part.xml) {
                let value = memberData.slice(part.start, part.end)
                str += "<w:r><w:t>" + value + "</w:t></w:r>";
            } else {
                let xmlObject = part.xml
                let curSegmentIdStr = _this.commentCount.toString()
                str += "<w:commentRangeStart w:id=\"" + curSegmentIdStr + "\"/>"
                str += "<w:r><w:t>"
                str += xmlObject.value
                str += "</w:t></w:r>"
                str += "<w:commentRangeEnd w:id=\"" + curSegmentIdStr + "\"/>"
                str += "<w:r><w:commentReference w:id=\"" + curSegmentIdStr + "\"/></w:r>"
                _this.comments[_this.commentCount++] = xmlObject.comment_str
            }
            return str;
        }, "")

        return result
    }

    recursive(options) {
        function getExpand(item) {
            return item.subparsed
        }

        function canExpand(item) {
            return !!getExpand(item)
        }

        let part = options.part;
        let data = options.data;

        if (!data)
            return;

        let need_move = []

        const result = part.reduce((parsed, item, index) => {
            if (item.type == "placeholder") {
                let node_name = item.value;
                let offset = 0;

                if (canExpand(item)) {
                    let data_value = data[node_name];

                    // 递归处理
                    this.recursive({
                        part: getExpand(item),
                        data: data_value
                    });

                } else if (this.isPlainText(item)) {
                    let hasXml = false;

                    if (data.length) {
                        // 是loop需要处理的数据
                        for (let i=0; i<data.length; ++i) {
                            let idxData = data[i]
                            const memberData = idxData[node_name]

                            // 判断是否含有xml标签
                            if (this.hasXmlObject(memberData)) {
                                hasXml = true;
                                break;
                            }
                        }

                        if (hasXml) {
                            for (let i=0; i<data.length; ++i) {
                                let idxData = data[i]
                                const memberData = idxData[node_name]
                                idxData[node_name] = this.convertToRawXml(memberData);
                            }
                        }
                    } else {
                        if (this.hasXmlObject(data)) {
                            data[node_name] = this.convertToRawXml(data)
                        }
                    }

                    if (hasXml) {
                        // 修改当前placeholder为rawxml类型
                        item.module = "rawxml"

                        if (index > 0) {
                            if (part[index - 1].type != "tag") {
                                // 这里假设，前一个不是tag元素，那当前placeholder前面可能还有其它元素，需要往右移动它
                                let endIndex = getRight(part, "w:r", index)

                                need_move.push({
                                    type: "move_right",
                                    from: item,
                                    to: part[endIndex]
                                });

                            } else if (index + 1 < part.length) {
                                // placeholder是tag标签中的第一个元素，并且不是最后一个元素
                                if (part[index + 1].type != "tag") 
                                {
                                    let beginIndex = getLeft(part, "w:r", index)

                                    need_move.push({
                                        type: "move_left",
                                        from: item,
                                        to: part[beginIndex]
                                    });
                                } 
                                else 
                                // 只有placeholder元素
                                {
                                    need_move.push({
                                        type: "expand",
                                        item,
                                    });
                                    //if (part[index-1].tag == "w:t") {

                                    /*} else*/ {
                                    }
                                }
                            }
                            else {
                                // 假设最后一个元素不能为placeholder
                                throwMalformedXml("placeholder can not be last position.")
                            }
                        } else {
                            // index == 0 假设第一个元素不能为placeholder
                            throwMalformedXml("placeholder can not be first position.")
                        }
                    }
                }
            }

            parsed.push(item);
            return parsed;
        }, []);

        function getInner({ part, left, right, postparsed, index }) {
            const before = getNearestLeft(postparsed, ["w:r"], left - 1);
            const after = getNearestRight(postparsed, ["w:r"], right + 1);
            //if (after === "w:tc" && before === "w:tc") {
            //part.emptyValue = "<w:p></w:p>";
            //}
            const paragraphParts = postparsed.slice(left + 1, right);
            paragraphParts.forEach(function(p, i) {
                if (i === index - left - 1) {
                    return;
                }
                if (isContent(p)) {
                    throwRawTagShouldBeOnlyTextInParagraph({ paragraphParts, part });
                }
            });
            return part;
        }

        // 移动带有comment的标记元素
        need_move.forEach(function(move_obj) {
            if (move_obj.type == "move_left") {
                let from_idx = part.indexOf(move_obj.from)
                let remove_item_list = part.splice(from_idx, 1)
                let to_idx = part.indexOf(move_obj.to)

                part.splice(to_idx, 0, remove_item_list[0])
            } else if (move_obj.type == "move_right") {
                let from_idx = part.indexOf(move_obj.from)
                let remove_item_list = part.splice(from_idx, 1)
                let to_idx = part.indexOf(move_obj.to)

                part.splice(to_idx + 1, 0, remove_item_list[0])
            } else if (move_obj.type == "expand") {
                let result = traits.expandOne(move_obj.item, part, {expandTo: "w:r", getInner});
                let args = concatArrays([[0, part.length], result])

                part.splice.apply(part, args);
            } else {
                throwMalformedXml("move_obj.type is unknown.")
            }
        });

        return result;
    }

    set(obj) {
        if (obj.xmlDocuments) {
            this.xmlDocuments = obj.xmlDocuments
        }

        if (obj.inspect && obj.inspect.postparsed) {
            this.commentCount = 0
            // 对postparsed后的数据做处理
            let postparsed = obj.inspect.postparsed
            let new_postparsed = this.recursive({
                    part: postparsed, 
                    data: this.data
                });

            //new_postparsed = this.rawxml.postparse(postparsed);
            let xml = this.generate_comments_xml()
            if (xml) {
                this.xmlDocuments["word/comments.xml"] = xml
            }
        }
    }

    hasXmlObject(data) {
        return re.exec(data) != null
    }

    getXmlObject(data) {
        // object
        // {
        //  start: number,
        //  end: number,
        //  type: string
        //  attribute: map<string,string>
        //  value: string
        // }
        
        let arr = null;
        let result = []
        let re = /\<comment\s+value=['"]([\w\t]+)['"]\>([^<]+)\<\/comment\>/g

        while ((arr = re.exec(data)) != null) {
            let start = arr.index;
            let end = start + arr[0].length
            let comment_str = arr[1]
            let value = arr[2]

            result.push(
                {
                    start,
                    end,
                    value,
                    comment_str,
                    type: "comment",
                }
            )
        }

        return result;
    }

    isPlainText(item) {
        return !item.module
    }

    compileToOxml(item) {
        return null
    }

    generate_comments_xml() {
        if (this.comments.length == 0)
            return 

        let doc = null;
        let comment_file = this.docxtemplater.zip.files["word/comments.xml"]
        let element_comments = null;

        if (comment_file) {
            // 判断是否已经存在comments.xml
            const usedData = comment_file.asText();
            doc = new DOMParser().parseFromString(usedData, 'text/xml');
            let element_comments_list = doc.getElementsByTagName("w:comments")

            if (element_comments_list) {
                element_comments = element_comments_list[0]
                let element_comment_list = element_comments.getElementsByTagName("w:comment")
                if (element_comment_list) {
                    for (let i=0; i<element_comment_list.length; ++i) {
                        let el = element_comment_list[i]
                        let comment_id = el.getAttribute("w:id")
                        if (comment_id && parseInt(comment_id) > this.commentCount) {
                           this.commentCount =  parseInt(comment_id) + 1
                        }
                    };
                }
            }
        } else {
            let example = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            doc = new DOMParser().parseFromString(example, 'text/xml');
        }

        if (!element_comments) {
            element_comments = doc.createElement("w:comments")
            element_comments.setAttribute("xmlns:wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas")
            element_comments.setAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
            element_comments.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office")
            element_comments.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            element_comments.setAttribute("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
            element_comments.setAttribute("xmlns:v", "urn:schemas-microsoft-com:vml")
            element_comments.setAttribute("xmlns:wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
            element_comments.setAttribute("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
            element_comments.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            element_comments.setAttribute("xmlns:w14", "http://schemas.microsoft.com/office/word/2010/wordml")
            element_comments.setAttribute("xmlns:w10", "urn:schemas-microsoft-com:office:word")
            element_comments.setAttribute("xmlns:w15", "http://schemas.microsoft.com/office/word/2012/wordml")
            element_comments.setAttribute("xmlns:wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
            element_comments.setAttribute("xmlns:wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk")
            element_comments.setAttribute("xmlns:wne", "http://schemas.microsoft.com/office/word/2006/wordml")
            element_comments.setAttribute("xmlns:wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
            element_comments.setAttribute("mc:Ignorable", "w14 w15 wp14")
            doc.appendChild(element_comments)
        }

        function generate_comment(options) {
            let element_comment = doc.createElement("w:comment")
            // comment id
            element_comment.setAttribute("w:id", options.id)
            // 作者
            element_comment.setAttribute("w:author", "作者")
            // 时间
            element_comment.setAttribute("w:date", "2019-08-20T10:04:31Z")
            // 
            element_comment.setAttribute("w:initials", "o")
            // 在w:comments标签加入w:comment标签
            element_comments.appendChild(element_comment)

            let element_w_p = doc.createElement("w:p")
            let element_w_pPr = doc.createElement("w:pPr")
            element_w_p.appendChild(element_w_pPr)
            let element_w_r = doc.createElement("w:r")
            element_w_p.appendChild(element_w_r)
            element_comment.appendChild(element_w_p)

            let element_w_pStyle = doc.createElement("w:pStyle")
            element_w_pStyle.setAttribute("w:val", "2")
            element_w_pPr.appendChild(element_w_pStyle)

            let element_w_rPr = doc.createElement("w:rPr")
            let element_w_rFonts = doc.createElement("w:rFonts")
            element_w_rFonts.setAttribute("w:hint", "eastAsia")
            element_w_rFonts.setAttribute("w:eastAsia", "宋体")
            element_w_rPr.appendChild(element_w_rFonts)
            let element_w_lang = doc.createElement("w:lang")
            element_w_lang.setAttribute("w:val", "en-US")
            element_w_lang.setAttribute("w:eastAsia", "zh-CN")
            element_w_rPr.appendChild(element_w_lang)
            element_w_pPr.appendChild(element_w_rPr)

            element_w_r.appendChild(element_w_rPr.cloneNode(true))
            let element_w_t = doc.createElement("w:t")
            // 评论内容
            element_w_t.textContent = options.comment_str
            element_w_r.appendChild(element_w_t)
        }

        for (let i in this.comments) {
            let options = {
                id: i, 
                comment_str: this.comments[i]
            }
            generate_comment(options)
        }

        //console.log(example + new XMLSerializer().serializeToString(element_comments))

        return doc
    }
}

module.exports = () => wrapper(new AddCommentModule());
