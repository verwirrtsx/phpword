<?php
namespace app\index\controller;

use think\Controller;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\Common\Font;

class Index extends Controller
{
    public function index()
    {
        // Create a new PHPWord Object
        $PHPWord = new PhpWord();
        $PHPWordHelper= new Font();

        $PHPWord->setDefaultFontName('隶书');  // 全局字体
        $PHPWord->setDefaultFontSize(16);       // 全局字号为3号


        // 添加3号仿宋字体到'FangSong16pt'留着下面使用
        $PHPWord->addFontStyle('FangSong16pt', array('name'=>'仿宋', 'size'=>16));

        // 添加段落样式到'Normal'以备下面使用
        $PHPWord->addParagraphStyle(
            'Normal',array(
                'align'=>'both',
                'spaceBefore' => 0,
                'spaceAfter' => 0,
                'spacing'=>$PHPWordHelper->pointSizeToTwips(2.8),
                'lineHeight' => 1.19,  // 行间距
                'indentation' => array( // 首行缩进
                    'firstLine' => $PHPWordHelper->pointSizeToTwips(32)
                )
            )
        );

        // Section样式：上3.5厘米、下3.8厘米、左3厘米、右3厘米，页脚3厘米
        // 注意这里厘米(centimeter)要转换为twips单位
        $sectionStyle = array(
            'orientation' => 'landscape',                              //页面方向： 默认竖向：null/横向：landscape
            'marginLeft' => $PHPWordHelper->centimeterSizeToTwips(0),
            'marginRight' => $PHPWordHelper->centimeterSizeToTwips(0),
            'marginTop' => $PHPWordHelper->centimeterSizeToTwips(0),
            'marginBottom' => $PHPWordHelper->centimeterSizeToTwips(0),
            'pageNumberingStart' => 1, // 页码从1开始
            'footerHeight' => $PHPWordHelper->centimeterSizeToTwips(0),
        );
        $section = $PHPWord->addSection($sectionStyle); // 添加一节

        // 下面这句是输入文档内容，注意这里用到了刚才我们添加的
        // 字体样式FangSong16pt和段落样式Normal
        $section->addText('某某某文案策划', 'FangSong16pt', 'Normal');
        $section->addLink(
            'https://blog.csdn.net/qq_28761593/article/details/86598754',
            'tp5+phpword',
            array(
                'size'=>20,
                'name'=>'微软雅黑',
                'bold'=>true,
                'Color'=>'#f00',),
            null
        );
        $section->addTextBreak(1); // 新起一个空白段落
        header("Content-Type: application/doc");
        header("Content-Disposition: attachment; filename=".date("YmdHis").".doc");
        $objWriter = IOFactory::createWriter($PHPWord, 'Word2007');

        $objWriter->save('php://output');
    }



}
