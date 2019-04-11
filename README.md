# email2PDFPortfolio
Uses VBA to read inbox emails, convert to PDF through word, add all attachments from the email to the pdf and save in windows/temp. Then import that pdf into an Adobe pdf portfolio file in a specified network share location through a trusted JS function (E2PDF.JS). Identifying the appropriate directory by matching file numbers from the email subject to folder titles in specified network share location. After importing the pdf to the portfolio it sets the field values, with another trusted JS function (E2PDF.JS), in the portfolio for the newly imported object, just like Adobe pdfMaker but for the entire inbox. It is also compatible with pdfMaker so manual entry is still possible for those emails without identifying information.
