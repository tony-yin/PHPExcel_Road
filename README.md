# PHPExcel_Road

---

> 最近工作用到PHPExcel这个excel做导出导入较多，碰到了一些大大小小的问题，写出来与大家分享，存在的一些不足和遗漏欢迎大家指正和补充:s

### 一、基础：小试牛刀

#### 1. 引用文件
```
yourpath . /phpexcel/PHPExcel.php
```

#### 2. 实例化phpexcel类
```
"xcel = new PHPExcel();
```

#### 3. 获取当前单sheet（多sheet会在下面讲）
```
$objexcel = "xcel->getActiveSheet();
```

##### 4. 合并单元格
```
$objexcel->mergeCells('A1:M1');
```

#### 5. 获取一个cell的样式
```
$objexcel->getStyle('A1');
```
+ 获取一个cell的字体样式
```
$cellFont = $objexcel->getStyle('A1')->getFont();
```
+ 设置字体大小
```
$fontStyle->setSize(15);
```
+ 设置字体是否加粗
```
$fontStyle->setBold(true);
```
+ 设置字体颜色
```
$fontStyle->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
```
+ 获取一行样式
```
$rowStyle = $objexcel->getStyle(1)->getRowDimension();
```
+ 设置行高度
```
$rowStyle->setRowHeight(2);
```
+ 获取一列样式
```
$columnStyle = $objexcel->getStyle('A')->getColumnDimension();
```
+ 设置列宽度
```
$columnStyle->setWidth(10);
```
+ 获取一列对齐样式
```
$alignStyle = $objexcel->getStyle('A')->getAlignment();
```
+ 设置水平居中：同一水平线上居中，即为左右的中间
```
$alignStyle>setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
```
+ 设置垂直居中：同一垂直线居中，即为上下的中间
```
$alignStyle->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
```
+ 自动换行
```
$$alignStyle->setWrapText(true);
```

#### 6. 获取指定版本excel写对象
如需更早的版本可将`Excel2007`换成`Excel5`
```
$write = PHPExcel_IOFactory::createWriter("xcel, 'Excel2007');
```

### 二、进阶：一些有用的小知识

#### 1.行列数字索引方法
> phpexcel一般获取cell或者获取列都是通过ABC这样的英文字母获取的，它也可以通过0、1、2、3这样的数字表示sheet中的列，从0开始，0对应A，1对应B，基本上大多数方法都是数字行列索引，例如getStyleByColumnAndRow($col,$row),默认列参数在前，行参数在后，更多的可以参加phpexcel源码；

#### 2. 单行或单列参数格式
> 有的时候一个方法需要行列两个参数，例如只需要某一行参数可写成(null, $row),例如只需要获得某一列参数可写成($col, null)

#### 3. 列的数字索引格式和字母索引格式互转
+ 数字转字符串
```php
PHPExcel_Cell::columnIndexFromString('A');  // Return 1 not 0;
```
+ 字符串转数字
```php
PHPExcel_Cell::stringFromColumnIndex(0);    // Return 'A';
```
#### 4.PHPExcel读取数字类型 
> PHPExcel读取的cell数字，类型都是double型，可用gettyle()方法检测类型，当初我一直使用is_int()方法无果，搞得焦头烂额。。。

#### 5. 多cell边框线设置
PHPExcel生成的表格如果你不加处理，是不会帮你生成边框线的，生成边框线的方法如下：
```php
$borderArray = array(
    'borders' => array(
        'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN
        )
    )
);
$objexcel->getStyle($col1, $row1,$col2, $row2)->applyFromArray($borderArray);
```
>注：
> 1. `getStyle()`可以看需求改为`getStyleByColumnAndRow()`方法通过数字行列索引读取style；
> 2. array中`PHPExcel_Style_Border::`后面有三种格式分别是`BORDER_THIN`和`BORDR_MEDIUM`，表示边框线的粗细；
> 3. `getStyle()`中的索引可以是静态的，也可以是动态的，一般是在导出excel的数据set完毕后填写左上角的单元格行列索引和右下角的单元格行列索引；

> 参考资料
>
> http://phpexcel.codeplex.com/workitem/22160
> http://phpexcel.codeplex.com/workitem/20150

#### 6. 多cell字体加粗处理
```
$objexcel->getStyle($pCoordinate)->applyFromArray(array(
    'font' => array(  
        'bold' => true,                                              
    ),                                                               
));
```

#### 7. 多cell字体颜色处理
```
$objexcel->getStyle($pCoordinate)->applyFromArray(array(
    'font' => array(
        'color' => array(
            'rgb' => 'ff0000',
        ),
    ),
));
```

#### 8. 多sheet导入
动态为当前sheet设置索引，然后获取当前sheet，便可循环读取每一个sheet内容
```php
$objexcel->setActiveSheetIndex($index);   //$index = 0 1 2 3
$objexcel->getActiveSheet();    //return sheet1 sheet2 sheet 3
```

#### 9. 固定格式excel读取在写入
> 当需求是给定一个一个模板excel，需要往里面塞数据，我们不一定要通过代码给它设定样式，如果这个模板变化不大，我们完全可以存放一个格式相同的静态文件，然后通过PHPExcel读取，再往里面塞数据，最后进行保存操作，可以达到一样的效果，并且可以节省大量的资源。

#### 10. 合并单元格导入问题
> 在特殊的表格中，合并单元格普遍存在，而多个单元格合并成的一个单元格，只能`setValue()`一次，而我们如何判断合并单元格的具体行列呢？
```php
$range = $start_cell->getMergeRange();  // 通过合并单元格的开始单元格比如‘A1’，获取合并范围‘A1:A4’
$cell->isInRange($range);    // 遍历之后每一个单元格便可通过isInRange()方法判断当前单元格是否在合并范围内
```
### 三、高级：特殊场景特殊手段
#### 1. 单元格文本格式数据处理 
> 一般excel单元格中数据的格式为数据类型，而`PHPExcel`中的`getValue()`方法读取的也是数据类型，当把数据从数据类型改为文本类型后，在`PHPExcel`中读出来的是`PHPExcel_RichText`类型，`getValue()`读取返回`PHPExcel_RichText`是一个`object`类型（`PHPExcel_RichText`数据保存格式）；那如何读取这一类的数据呢？仔细查看读取出来的对象，不难发现有`getPlainText()`这样的方法可以读取文本类型数据，所以我们只要判断当当前数据为文本数据时用`getPlainText()`读取，一般数据用`getValue()`读取
```php
if ($cell->getValue() instanceof PHPExcel_RichText) {
    $value = $cell->getValue();
} else {
    $value = $cell->getValue();
}
```
> 参考资料
>
> http://www.cnblogs.com/DS-CzY/p/4955655.html
> http://phpexcel.codeplex.com/discussions/34513

#### 2. 单元格数据算法处理
> excel拥有强大的算法功能，一般算法格式为`=A3+A4`这类的，复杂的更多，如果使用PHPExcel提供的默认读取方法`getValue()`读取出来的结果则为字符串'=A3+A4',好在PHPExcel也足够强大，提供了相应的接口：`getCalculatedValue()`，这个方法专门读取算法数据，但是我们不能将这个方法作为默认读取方法，因为这样可能会将一些本来要读成字符串的读成算法数据，而且PHPExcel没有将它作为默认读取方法的另一个重要原因就是算法方式读取很耗时间和性能，一般数据读取根本没有必要这样浪费资源，所以我们可以采用以下这种方式
```php
if (strstr($cell->getValue(), '=')) {   
    // 判断如果cell内容以=号开头便默认为算法数据
    $value = $cell->getCalculatedValue(); 
} else {
    $value = $cell->getValue();
}
```

#### 3. 日期数据处理
> 除了以上所说的文本数据和算法数据外，我还遇到过日期类型数据，比如2016-12-28输入到excel中，它会默认转换成2016/12/28，如果采用一般的`getValue()`方式读取也会读取到错误的数据，PHPExcel也提供了相应的接口`getFormattedValue()`,并提供了适配的识别方式`PHPExcel_Shared_Date::isDateTime($cell)`,所以代码就很好实现了
```php
if (PHPExcel_Shared_Date::isDateTime($cell)) {
    $value = $cell->getFormattedValue(); 
} else {
    $value = $cell->getValue();
}
```
#### 4. 读取方法封装
> 针对excel各种数据类型，我们可以写一个函数，将原有的`getValue()`封装一下，这样以后就不用每次都判别一下数据类型了，目前我只遇到上面三种特殊格式，如果有新的，欢迎大家补充，封装函数如下
```php
function get_value_of_cell($cell) {
    if (strstr($cell->getValue(), '=')) {   
        $value = $cell->getCalculatedValue(); 
    } else if ($cell->getValue() instanceof PHPExcel_RichText) {
        $value = $cell->getValue();
    } else if (PHPExcel_Shared_Date::isDateTime($cell)) {
        $value = $cell->getFormattedValue(); 
    } else {
        $value = $cell->getValue();
    }
}
```
#### 5. 导出文件在IE、360等浏览器中文件名中文乱码问题
```php
$filename = 'xxx导出表';
// 判断如果是IE内核形式的浏览器采用urlencode处理文件名
if (!preg_match("/Firefox/", $_SERVER["HTTP_USER_AGENT"])) {
    $filename = urlencode($filename);
}
```
