function ConvertGoogleDocToCleanHtml() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var output = [];
  var images = [];
  var listCounters = {};

  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    output.push(processItem(child, listCounters, images));
  }

  var html = output.join('\r');
  emailHtml(html, images);
  //createDocumentForHtml(html, images);
}

function emailHtml(html, images) {
  var attachments = [];
  for (var j=0; j<images.length; j++) {
    attachments.push( {
      "fileName": images[j].name,
      "mimeType": images[j].type,
      "content": images[j].blob.getBytes() } );
  }

  var inlineImages = {};
  for (var j=0; j<images.length; j++) {
    inlineImages[[images[j].name]] = images[j].blob;
  }

  var name = DocumentApp.getActiveDocument().getName()+".html";
  attachments.push({"fileName":name, "mimeType": "text/html", "content": html});
  MailApp.sendEmail({
     to: Session.getActiveUser().getEmail(),
     subject: name,
     htmlBody: html,
     inlineImages: inlineImages,
     attachments: attachments
   });
}

function createDocumentForHtml(html, images) {
  var name = DocumentApp.getActiveDocument().getName()+".html";
  var newDoc = DocumentApp.create(name);
  newDoc.getBody().setText(html);
  for(var j=0; j < images.length; j++)
    newDoc.getBody().appendImage(images[j].blob);
  newDoc.saveAndClose();
}

function dumpAttributes(atts) {
  // Log the paragraph attributes.
  for (var att in atts) {
    Logger.log(att + ":" + atts[att]);
  }
}

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "", suffix = "";
  // var textElements = [];
 
  // Punt on TOC.
  if (item.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
      return {"text": "[[TOC]]"};
  }
  
  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING6: 
        prefix = "<h6>", suffix = "</h6>"; break;
      case DocumentApp.ParagraphHeading.HEADING5: 
        prefix = "<h5>", suffix = "</h5>"; break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>", suffix = "</h4>"; break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>", suffix = "</h3>"; break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>", suffix = "</h2>"; break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>", suffix = "</h1>"; break;
      default: 
        prefix = "<p>", suffix = "</p>";
    }

    if (item.getNumChildren() == 0)
      return "";
  }
  else if (item.getType()===DocumentApp.ElementType.LIST_ITEM) {
    var listItem = item;
    var gt = listItem.getGlyphType();
    var key = listItem.getListId() + '.' + listItem.getNestingLevel();
    var counter = listCounters[key] || 0;

    // First list item
    if ( counter == 0 ) {
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix = '<ul><li><p>', suffix = "</p></li>";

          suffix += "</ul>";
        }
      else {
        // Ordered list (<ol>):
        prefix = "<ol><li><p>", suffix = "</p></li>";
      }
    }
    else {
      prefix = "<li><p>";
      suffix = "</p></li>";
    }

    if (item.isAtDocumentEnd() || item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM) {
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        suffix += "</ul>";
      }
      else {
        // Ordered list (<ol>):
        suffix += "</ol>";
      }

    }

    counter++;
    listCounters[key] = counter;
  }

  output.push(prefix);
  
  if (item.getType() === DocumentApp.ElementType.TABLE)
  {
      processTable(item, output);
  }else if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output);
  }
  else {


    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();

      // Walk through all the child elements of the doc.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child, listCounters, images));
      }
    }

  }

  output.push(suffix);
  return output.join('');
}

function getAllLinks(element) {
  var links = [];
  element = element || DocumentApp.getActiveDocument().getBody();
  
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var textObj = element.editAsText();
    var text = element.getText();
    var inUrl = false;
    for (var ch=0; ch < text.length; ch++) {
      var url = textObj.getLinkUrl(ch);
      if (url != null) {
        if (!inUrl) {
          // We are now!
          inUrl = true;
          var curUrl = {};
          curUrl.element = element;
          curUrl.url = String( url ); // grab a copy
          curUrl.startOffset = ch;
        }
        else {
          curUrl.endOffsetInclusive = ch;
        }          
      }
      else {
        if (inUrl) {
          // Not any more, we're not.
          inUrl = false;
          links.push(curUrl);  // add to links
          curUrl = {};
        }
      }
    }
  }
  else {
    // Get number of child elements, for elements that can have child elements. 
    try {
      var numChildren = element.getNumChildren();
    }
    catch (e) {
      numChildren = 0;
    }
    for (var i=0; i<numChildren; i++) {
      links = links.concat(getAllLinks(element.getChild(i)));
    }
  }

  return links;
}

function processText(item, output) {
  if (!item) return;
 
  var text = item.getText();
  var indices = item.getTextAttributeIndices();
  var links = getAllLinks();
 
  if (indices.length <= 1) {
    // Assuming that a whole para fully italic is a quote
    if(item.isBold()) {
      output.push('<b>' + text + '</b>');
    }
    else if(item.isItalic()) {
      output.push('<blockquote>' + text + '</blockquote>');
    }
    else if (text.trim().indexOf('http://') == 0) {
      output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
    }
    else {
      output.push(text);
    }
  }
  else {

    for (var i=0; i < indices.length; i ++) {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i+1 < indices.length ? indices[i+1]: text.length;
      var partText = text.substring(startPos, endPos);
      var link = links[i];
      
      Logger.log(partText);

      if (partAtts.ITALIC) {
        output.push('<i>');
      }
      if (partAtts.BOLD) {
        output.push('<b>');
      }
      if (partAtts.UNDERLINE) {
        output.push('<a href="' + link.url + '" rel="nofollow"' + 'title= "'+ partText + '" class= "'+ partText + '" target="_blank"' + '>' + partText + '</a>');
          // output.push('<u>');
     }
     if (partAtts.EMPHASIS){
        output.push('<em>');
      }


      // If someone has written [xxx] and made this whole text some special font, like superscript
      // then treat it as a reference and make it superscript.
      // Unfortunately in Google Docs, there's no way to detect superscript
      if (partText.indexOf('[')==0 && partText[partText.length-1] == ']') {
        output.push('<sup>' + partText + '</sup>');
      }
      else if (partText.trim().indexOf('http://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      }
      else {
        output.push(partText);
      }

      if (partAtts.ITALIC) {
        output.push('</i>');
      }
      if (partAtts.BOLD) {
        output.push('</b>');
      }
      if (partAtts.UNDERLINE) {
        output.push('</u>');
      }
      if (partAtts.EMPHASIS){
        output.push('<em>');
      }

    }
  }
}

function processTable(item, output) {
  if (!item) return;
  if (item.getType() === DocumentApp.ElementType.TABLE) {
      output.push("<table>\n");
      output.push("<thead>\n");
      var nCols = item.getChild(0).getNumCells();
      for (var i = 0; i < item.getNumChildren(); i++) {
        output.push("  <tr>\n");
        // process this row
        for (var j = 0; j < nCols; j++) {
            if (item.getChild(i).getChild(j).getBackgroundColor()){

              if (item.getChild(i).getChild(j).isBold() ){
                output.push("<td style='text-align:" + item.getRow(i).getCell(j).getChild(0).getAlignment() + ';' + 'background-color:' + item.getRow(i).getCell(j).getBackgroundColor() + "'>"  + "<strong>" + item.getChild(i).getChild(j).getText()+ "</strong>" + "</td>\n");
              } else if (item.getChild(i).getChild(j).isItalic() ) {
                output.push("<td style='text-align:" + item.getRow(i).getCell(j).getChild(0).getAlignment() + ';' + 'background-color:' + item.getRow(i).getCell(j).getBackgroundColor() + "'>"  + "<em>" + item.getChild(i).getChild(j).getText()+ "</em>" + "</td>\n");
              }
              else {
                output.push("<td style='text-align:" + item.getRow(i).getCell(j).getChild(0).getAlignment() + ';' + 'background-color:' + item.getRow(i).getCell(j).getBackgroundColor() + "'>"  +  item.getChild(i).getChild(j).getText()+ "</td>\n");
              }
            }
          else {
            if (item.getChild(i).getChild(j).isBold() ){
              output.push("<td style='text-align:" + item.getRow(i).getCell(j).getChild(0).getAlignment() + "'>"  + "<strong>" + item.getChild(i).getChild(j).getText() + "</strong>" +"</td>\n");
            }
            else if (item.getChild(i).getChild(j).isItalic()){
              output.push("<td style='text-align:" + item.getRow(i).getCell(j).getChild(0).getAlignment() + "'>"  + "<em>" + item.getChild(i).getChild(j).getText() + "</em>" +"</td>\n");
            }
            else {
              output.push("<td style='text-align:" + item.getRow(i).getCell(j).getChild(0).getAlignment() + "'>"  + item.getChild(i).getChild(j).getText() + "</td>\n");
            }
          }
        }
        output.push("  </tr>\n");
      }
      output.push("</tbody>\n");
      output.push("</table>\n");
  }
}

