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


