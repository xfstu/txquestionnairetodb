<?php

namespace app\index\controller;

use xfstu\tools\ToolsXlsx;

class Index
{
    public function test()
    {
        $files = scandir('xlsx');
        $files = array_diff($files, ['.', '..']);
        $filesPath = array_map(function ($item) {
            return ['xlsx/' . $item, pathinfo($item)['filename']];
        }, $files);
        $obj = new ToolsXlsx();
        $obj->getFileArray('user.xlsx', 2, function ($row) use ($filesPath, $obj) {
            foreach ($filesPath as $key => $value) {
                if ($row[3] == $value[1]) {
                    $tid = $value[1];
                    $obj->getFileArray($value[0], 2, function ($row) use ($tid) {
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

    }
}
