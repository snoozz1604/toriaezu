"toriaezu" 

# PHPExcel Memo

### ループ処理１） foreach × 2

    <?php
    $excel = new PHPExcel;
    $sheet = $excel->getActiveSheet();
    $sheet->setTitle('ジョジョシート');
    
    $data_set = array(
      array('id' => 1, 'name' => 'ジョナサン', 'stand' => null),
      array('id' => 2, 'name' => 'ジョセフ', 'stand' => 'ハーミット・パープル'),
      array('id' => 3, 'name' => '空条承太郎', 'stand' => 'スター・プラチナ'),
      array('id' => 4, 'name' => '東方仗助', 'stand' => 'クレイジー・ダイヤモンド'),
      array('id' => 5, 'name' => 'ジョルノ', 'stand' => 'ゴールド・エクスペリエンス'),
      array('id' => 6, 'name' => '空条徐倫', 'stand' => 'ストーン・フリー'),
      array('id' => 7, 'name' => 'ジョニィ', 'stand' => 'タスク'),
      array('id' => 8, 'name' => '東方定助', 'stand' => 'ソフト＆ウェット')
    );
    
    
    $row = 1;
    foreach ($data_set as $data) {
      $col = 0;
    
      foreach ($data as $value) {
        $sheet->setCellValueByColumnAndRow($col++, $row, $value);
      }
      $row++;
    }
    
    $writer = PHPExcel_IOFactory::createWriter($excel, 'Excel5');
    $writer->save('/tmp/jojo.xls');


### ループ処理２） fromArray メソッド

    <?php
    $sheet->fromArray($data_set, null, 'A1');


### スタイルのコピー）duplicateStyle

文字通りスタイルの複製。あるエリアのスタイルをコピーして使う。

    $sheet->getStyleByColumnAndRow($col,$row);
    $sheet->duplicateStyle（$style,'A1');

### セルの値（数式も含む）のコピー） setCellValue ＋ >getValue

    $sheet->getCellByColumnAndRow($col,$row);
    $sheet->setCellValue('A1',$cell->getValue());

この$cellはコピー元のセルオブジェクト。コピー元のセルの値を取得して、そのまま新しいセルの値としてセット。


### R1C1 ⇒ A1形式

１）セルの番地が欲しい場合（R1C1をA1に変更）… 'stringFromColumnIndex'

    $alfa_col = PHPExcel_Cell::stringFromColumnIndex($col);

　カラムは0から番号が振ってあるので、1がAではなく、0がA。

　
２）計算式が入ってるセルの結果だけが欲しい場合 … 　'getCalculatedValue’

    $sheet->getCellByColumnAndRow($col,$row)->getCalculatedValue();


３）ちなみに、セルに計算式を入れたい場合

    $sheet->setCellValueByColumnAndRow($col,$row,$string);

　セルに入れる計算式 $string は "=A1 / 10"とかでOK。

　
