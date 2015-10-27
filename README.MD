<h1><span style="font-size: 24pt;"><strong>功能介绍：</strong></span></h1>
<span style="font-size: 14pt;">本库可以将Excel数据反序列化成便于操作的对象。</span>
<h1><span style="font-size: 24pt;"><strong>实现原理：</strong></span></h1>
<img class="alignnone" src="http://7xiot4.com1.z0.glb.clouddn.com/ExcelDeserializer.png-blog" alt="" width="400" height="1094" />

<span style="font-size: 14pt;">首先用Poi读取文件，得到文件头，对应类的属性，然后再读取所有表格数据，构建Json字符串，再借助fastjson反序列化为对象。</span>

&nbsp;
<h1><span style="font-size: 24pt;"><strong>使用方法：</strong></span></h1>
Excel表格数据


<table>
<tbody>
<tr><td><em>name</em></td><td><em>srore </em></td><td><em>note</em></td></tr>
<tr><td>黄鹏</td><td>8</td><td></td></tr>
<tr><td>张辉</td><td>7</td><td></td></tr>
<tr><td>汪旭</td><td>9</td><td></td></tr>
<tr><td>黄祚</td><td>3</td><td>优秀</td></tr>
<tr><td>李阳</td><td>6</td><td></td></tr>
<tr><td>文凡</td><td>7</td><td></td></tr>
<tr><td>赵伟</td><td>3</td><td></td></tr>
<tr><td>梁琳</td><td>3</td><td></td></tr>


</tbody>
</table>




 
    public class TestModel {
    public String name;
    public String score;
    public String note;

    @Override
    public String toString() {
        return "TestModel{" +
                "name='" + name + '\'' +
                ", score='" + score + '\'' +
                ", note='" + note + '\'' +
                "}\n ";
    }
}


&nbsp;



    public class Test {
    public static void main(String[] args) {
        File file = new File(System.getProperty("user.dir") + "\\excel" + '\\' + "transcript.xls");
        ExcelDeserializer controller = new ExcelDeserializer();
        List<TestModel> models=controller.read(file, TestModel.class);
    }
}


<span style="font-size: 18pt;">结果</span>

[TestModel{name='黄鹏', score='8', note=''}

, TestModel{name='张辉', score='7', note=''}

, TestModel{name='汪旭', score='9', note='优秀'}

, TestModel{name='黄祚', score='3', note=''}

, TestModel{name='李阳', score='6', note=''}

, TestModel{name='文凡', score='7', note=''}

, TestModel{name='赵伟', score='3', note=''}

, TestModel{name='梁琳', score='3', note=''}
]