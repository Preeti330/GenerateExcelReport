<?php 
// query to get redemption report of customer 
// based o status give flag as scanned,activated, redemeed 
//pass the query result as array to generte excel report . 
//code is according to Yii2 FrameWork Stracture for genareting URL And Using Spreadsheet class to prepare Xlxs sheet.
//any query can passs to method " generateExcelReport()" to generate excel data .


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use yii\helpers\Url;

$conn = new mysqli('localhost', 'root', '','coke_and_meals');
//check DB connection 
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
  }

$sql="SELECT ct.id AS trancation_id, o.id AS combo_no,o.offer_name AS combo_name,o.discount_value AS combo_value,r.restaurant_code AS unquie_restaurant_id,r.restaurant_name,ct.customer_id,c.customer_name,c.mobile_no,c.pincode,states.state_name,c.email,c.city,o.status AS order_status,
CASE WHEN ct.status=2 THEN 'Offer_Activated'
WHEN ct.status=0 THEN 'Scanned'
WHEN ct.status=1 THEN 'Redemeed'
END AS reedmption_result,ct.created_date
FROM customer_transactions AS ct
INNER JOIN (SELECT customers.id,customers.customer_name,customers.mobile_no,customers.city,customers.address_lane_1,customers.pincode,customers.state_id,customers.email FROM customers) AS c 
ON(c.id=ct.customer_id)
JOIN restaurant_offers AS ro
ON(ro.id=ct.restaurant_offer_id)
JOIN offers AS o
ON(o.id=ro.offer_id)
JOIN restaurants AS r
ON(r.id=ro.restaurant_id)
JOIN states 
ON(states.id=r.state_id)  ";



$queryResultData=Yii::$app->db->createCommand($sql2)->queryAll();

$SpreadSheetTitle="RedeemptionReport";//sheetName 
$url=generateExcelReport($queryResultData, $SpreadSheetTitle);


function generateExcelReport($queryResultData, $SpreadSheetTitle)
{
    $spreadsheet = new Spreadsheet();
    $activesheet = $spreadsheet->getActiveSheet();
    //create file directory
    $dir = 'uploads/';
    if (!file_exists($dir)) {
        mkdir($dir, 077, true);
    }
    $date=date('Y-m-d');
    $dirpath = 'uploads/reports/';
    $file_name = uniqid() . '.Xls';
    $file_name = $dirpath .$date.'-'.$file_name;
    //set title of spreadsheet
    $activesheet->setTitle($SpreadSheetTitle);
    //set row as i=2;
    $i = 2;
    foreach ($queryResultData as $key => $val) 
    {
        //$count = count($queryResultData[$key]);
        $j = 0;
        foreach ($val as $keys => $value) {
            //assci val of 65 is A and Every Time increase it by one to get next column (like B,C.......)               
            $charval = 65 + $j;
             $char=chr($charval);                       
            //set headers using $key for $char.1(A1,B1....);
            if($i==2)
            {
                if($char>='A' && $char<='Z'){
                    $activesheet->setCellValue($char.'1',$keys);
                    //set style to headers 
                    $spreadsheet->getActiveSheet()->getStyle($char.'1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('green');
                    $spreadsheet->getActiveSheet()->getStyle($char.'1')->getFont()->setBold(true);
                }
            }
            if ($charval >= 65 && $charval <= 90) 
            {
                $char = chr($charval);
                //  set values remaining cloumns 
                $activesheet->setCellValue($char . $i, $val[$keys]);
            }
            if ($charval >= 91 && $charval<=116){
                //set values for Column AA,AB,AC,.....
                $charval2 = $charval - 26;//91-26=65(i.e A)
                $charA = 'A'.chr($charval2);
                if($i==2){

                    if($charA>='AA' && $charA<='AZ'){
                        $activesheet->setCellValue($charA.'1',$keys);
                        //set style to headers 
                        $spreadsheet->getActiveSheet()->getStyle($charA.'1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('green');
                        $spreadsheet->getActiveSheet()->getStyle($charA.'1')->getFont()->setBold(true);
                    }
                }
               // echo "charval $charval2 $charA <br>";
                 $activesheet->setCellValue($charA.$i,$val[$keys]);
            }

            if($charval>=117 && $charval<=142)
            {
                 //set values for Column BA,BB,BC,BD,.....
                $charval2 = $charval - 52;
                $charB = 'B'.chr($charval2);
                if($i==2){
                if($charB>='BA' && $charB<='BZ'){
                    $activesheet->setCellValue($charB.'1',$keys);
                    $spreadsheet->getActiveSheet()->getStyle($charB.'1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('green');
                    $spreadsheet->getActiveSheet()->getStyle($charB.'1')->getFont()->setBold(true);
                }
            }
                 $activesheet->setCellValue($charB.$i,$val[$keys]);
            }             
            //incriment j by one inside second loop to change char value 
            $j = $j + 1;
        }
        //increment i for new row
        $i++;
    }
    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xls');
    $writer->save($file_name);
    $url = Url::home(true) . $file_name;
    return $url;
}



?>