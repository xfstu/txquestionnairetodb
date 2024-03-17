# 目的

这是一个将腾讯问卷一个项目的所有问卷的数据合并在一个数据库中，以便使用程序分析所有问卷的数据。

# python

先建立一个python虚拟环境（推荐），这个步骤主要是将腾讯问卷导出的zip自动解压，将csv转为PHP容易读取的xlsx，顺便将文件名命名为问卷ID，方便后续处理数据。

```shell
python -m venv myenv
pip install pandas
```

把腾讯问卷导出的zip放到当前文件夹下，然后执行命令:`python.exe python.py`

# PHP

将python转换的csv放入web根目录xlsx，然后在你的控制器使用这段代码

```php
$files = scandir('xlsx');
$files = array_diff($files, ['.', '..']);
$filesPath = array_map(function ($item) {
    return ['xlsx/' . $item, pathinfo($item)['filename']];
}, $files);
$obj = new ToolsXlsx();
//以下代码逻辑自己定义
//user.xlsx主要是每个问卷的信息。主要逻辑是：一份问卷表，一份调查表，一个问卷表有一个调查表，一个调查表有多条信息，现在要做的就是循环读取问卷表，每一条信息对应一个调查表，再将调查表全部读取，并将问卷表ID（腾讯问卷ID）和每一条信息关联起来。
$obj->getFileArray('user.xlsx', 2, function ($row) use ($filesPath, $obj) {
    foreach ($filesPath as $key => $value) {
        if ($row[3] == $value[1]) {
            $tid = $value[1];
            $obj->getFileArray($value[0], 2, function ($row) use ($tid) {
                //这里的逻辑自定义
                $data = [
                    'start' =>  date('Y-m-d H:i', strtotime($row[1])),
                    'ida'   =>  $row[8],
                    'alias' =>  $row[9],
                    'a'     =>  getCode($row[14]),
                    'b'     =>  getCode($row[15]),
                    'c'     =>  getCode($row[16]),
                    'd'     =>  $row[17],
                    'e'     =>  $row[18],
                    'f'     =>  $row[19],
                    'tid'   =>  $tid,
                ];
                Db::table('share2023_table')->insert($data);
            });
        }
    }
});
```

`ToolsXlsx`类主要是将xlsx转为二维数组，需要的可以联系我。