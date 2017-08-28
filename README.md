# toExcel
将网页表格转变为excel文件   支持IE8,主流浏览器

使用:    
1.先引入js文件

    <script type="text/javascript" src="toExcel.js"></script>    
2.在表格外面设置一个容器包裹

        <div id="myDiv">  
        <table id="tableExcel" width="100%" border="1" cellspacing="0" cellpadding="0">  
        .....
        </table>
        </div>
        
        
3.使用toExcel

        var excel=new ToExcel({fileName:"LALALA.xls"});//设置文件名    
        excel.method("myDiv");//获取容器的div做为参数
        

