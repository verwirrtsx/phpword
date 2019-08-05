<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2019/8/5
 * Time: 9:39
 */

namespace app\index\controller;

use think\Controller;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\Common\Font;
use PhpOffice\PhpWord\IOFactory;

class Add extends Controller
{
    public  function  index(){
        $tmp=new \PhpOffice\PhpWord\TemplateProcessor('http://localhost/word/public/word/tmp.docx');//打开模板
        //dump($tmp);
        $tmp->setValue('name','李四');//替换变量name
        $tmp->setValue('mobile','18888888888');//替换变量mobile
        $tmp->setValue('school','社会大学');//替换变量mobile
        $tmp->saveAs('jianli.docx');//另存为
       // $tmp->save('php://output');
    }
}