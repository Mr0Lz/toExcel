       function ToExcel(opt){
        var idTmr;
        var a;  
        this.method=method;
        function  IE() { 
            if (!!window.ActiveXObject || "ActiveXObject" in window ) {  
                return 'ie';  
            }  
        }  
        function method(tableid) {  
            if(IE()=='ie')  
            {  
                var curTbl = document.getElementById(tableid);  
                var oXL = new ActiveXObject("Excel.Application");  
                var oWB = oXL.Workbooks.Add();  
                var xlsheet = oWB.Worksheets(1);  
                var sel = document.body.createTextRange();  
                sel.moveToElementText(curTbl);  
                sel.select();  
                sel.execCommand("Copy");  
                xlsheet.Paste();  
                oXL.Visible = true;  
  
                try {  
                    var fname = oXL.Application.GetSaveAsFilename(opt.fileName||"Excel.xls", "Excel Spreadsheets (*.xls), *.xls");  
                } catch (e) {  
                    print("Nested catch caught " + e);  
                } finally {  
                    oWB.SaveAs(fname);  
                    oWB.Close(savechanges = false);  
                    oXL.Quit();  
                    oXL = null;  
                    idTmr = window.setInterval("Cleanup();", 1);  
                }  
  
            }  
            else  
            {  
                tableToExcel(tableid,opt.fileName||"Excel.xls");
                CleanDom();
            }  
        }  
        function CleanDom(){
             document.body.removeChild(a);
        }
        function Cleanup() {  
            window.clearInterval(idTmr);  
        }
        function download(href,name){
            a=document.createElement("a");
            a.href=href;
            a.id="centent";
            a.download=name;
            document.body.appendChild(a);
            a.click();
        }  
        var tableToExcel = (function() {  
            var uri = 'data:application/vnd.ms-excel;base64,';
                    template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--><head><meta charset="UTF-8"></head><body><table border="1">{table}</table></body></html>',  
                    base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) },  
                    format = function(s, c) {
                        return s.replace(/{(\w+)}/g,function(m, p) { return c[p]; }) }  
            return function(table, name) {  
                if (!table.nodeType) table = document.getElementById(table)  
                var ctx = {worksheet: name, table: table.innerHTML}  
               download((uri + base64(format(template, ctx))),ctx.worksheet);
            }  
        })()  
    }