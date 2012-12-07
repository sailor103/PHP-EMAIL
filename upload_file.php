<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr" lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Excel邮件群发</title>
<link rel="stylesheet" type="text/css" media="all" href="style.css" />
</head>
<body>
<h1>Excel邮件群发</h1>
<!--外围框-->
<div id="outfrmB">
	<!--选择按钮-->
	<div id="selectbtn">
		<?php 
			set_time_limit(0);
			require_once('class.phpmailer.php');
		    require_once("class.smtp.php"); 


		function yqc_mail($to,$subject = "",$body = ""){
		    //Author:Jiucool WebSite: http://www.jiucool.com 
		    //$to 表示收件人地址 $subject 表示邮件标题 $body表示邮件正文
		    //error_reporting(E_ALL);
		    error_reporting(E_STRICT);
		    date_default_timezone_set("Asia/Shanghai");//设定时区东八区
		    
		    $mail             = new PHPMailer(); //new一个PHPMailer对象出来
		    $body             = eregi_replace("[\]",'',$body); //对邮件内容进行必要的过滤
		    $mail->CharSet ="UTF-8";//设定邮件编码，默认ISO-8859-1，如果发中文此项必须设置，否则乱码
		    $mail->IsSMTP(); // 设定使用SMTP服务
		    $mail->SMTPDebug  = 1;                     // 启用SMTP调试功能
		                                           // 1 = errors and messages
		                                           // 2 = messages only
		    $mail->SMTPAuth   = true;                  // 启用 SMTP 验证功能
		    $mail->SMTPSecure = "ssl";                 // 安全协议
		    $mail->Host       = "smtp.gmail.com";      // SMTP 服务器
		    $mail->Port       = 465;                   // SMTP服务器的端口号
		    $mail->Username   = "sailor9066";  // SMTP服务器用户名
		    $mail->Password   = "yinquanchao#";            // SMTP服务器密码
		    $mail->SetFrom("sailor9066@gmail.com", "cugxx");
		    $mail->AddReplyTo("sailor9066@gmail.com","cugxx");
		    $mail->Subject    = $subject;
		    $mail->AltBody    = "To view the message, please use an HTML compatible email viewer!"; // optional, comment out and test
		    $mail->IsHTML(true);
			$mail->MsgHTML($body);
		    $address = $to;
		    $mail->AddAddress($address,$to);
		    //$mail->AddAttachment("images/phpmailer.gif");      // attachment 
		    //$mail->AddAttachment("images/phpmailer_mini.gif"); // attachment
		    if(!$mail->Send()) {
		        echo "Mailer Error: " . $mail->ErrorInfo;
		    } else {
		        echo "邮件发送成功！<br/>";
		        }
		    }
		


		if (file_exists("upload/" . $_FILES["file"]["name"]))
		  {
		  //echo $_FILES["file"]["name"] ;
		  }
		else
		  {
		  move_uploaded_file($_FILES["file"]["tmp_name"],"upload/" . $_FILES["file"]["name"]);
		  //echo "Stored in: " . "upload/" . $_FILES["file"]["name"];
		  }
		  
		//开始读取这个文件
		set_include_path('.'. PATH_SEPARATOR .  
             'Classes' . PATH_SEPARATOR .  
             get_include_path());  

		require_once 'PHPExcel.php'; 
		
		$filePath="upload/".$_FILES["file"]["name"];//定义一个变量作为文档存储路径
		$PHPExcel = new PHPExcel();  
		$PHPReader = new PHPExcel_Reader_Excel2007();  
		//if exist the excel file
		if(!$PHPReader->canRead($filePath)){      
		 $PHPReader = new PHPExcel_Reader_Excel5(); 
		 if(!$PHPReader->canRead($filePath)){      
		  echo 'no Excel';
		  return ;
		 }
		}
		//read the file
		$PHPExcel = $PHPReader->load($filePath);
		$currentSheet = $PHPExcel->getSheet(0);
		$banji=$currentSheet->getTitle();
		//get info about mail
		for($row=2;$row<=13;$row++)
		{
			//邮件地址
			$mailto=$currentSheet->getCell('E'.$row)->getValue();
			echo $mailto;
			//主题
			$mailsub="自主创新团队第一次调研结果反馈";
			//邮件内容
			$mailcon='<h1>自主创新团队第一次调研结果反馈</h1>';
			$stuno=$currentSheet->getCell('B'.$row)->getValue();
			$mailcon=$mailcon.'<p>亲爱的'.$stuno.'同学：</p>';
			$mailcon=$mailcon.'<p style="text-indent: 2em">你好。自主创新团队的第一次调研共评估了参与创新团队同学们的认知风格和倾向。包括：突破创新倾向、责任意识、服从倾向、渐进创新倾向、积极进取意识、参加创新团队的内在动机和外在动机、团队倾向、创新效能感、团队中的自尊感、风险规避倾向、开放性、尽责性、掌握目标倾向、表现目标倾向、回避目标倾向、个人-集体偏好、权力距离意识。观点采择能力、以及实际的创新卷入行为。</p>';
			$mailcon=$mailcon.'<p style="text-indent: 2em">以下，我们将参与团队的人的得分情况以及你个人的得分情况进行比较。在表格中：</p>';
			$mailcon=$mailcon.'<p>25%表示：在所有参与调研的同学中，25%的同学的得分在这个分数以下</p>';
			$mailcon=$mailcon.'<p>50%表示：在所有参与调研的同学中，50%的同学的得分在这个分数以下，即平均分数</p>';
			$mailcon=$mailcon.'<p>75%表示：在所有参与调研的同学中，75%的同学的得分在这个分数以内。即只有25%的同学得分超过此分数</p>';
			$mailcon=$mailcon.'<p>在每个维度上得分越高，表明你在这个方面的倾向会越明显，越突出。</p>';
			//写入表格数据
			$mailcon=$mailcon.'<table  border="1">';
			//总共22行数据
			$mailcon=$mailcon.'<tr> <td></td> <td>25%</td> <td>50%</td> <td>75%</td> <td>你的得分</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>突破创新倾向</td> <td>18</td> <td>20.5</td> <td>23</td> <td>'.$currentSheet->getCell('R'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>责任意识</td> <td>21</td> <td>23</td> <td>25</td> <td>'.$currentSheet->getCell('S'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>服从倾向</td> <td>18</td> <td>21</td> <td>23</td> <td>'.$currentSheet->getCell('T'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>渐进创新倾向</td> <td>15</td> <td>17</td> <td>20</td> <td>'.$currentSheet->getCell('U'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>积极进取意识</td> <td>21</td> <td>23</td> <td>25</td> <td>'.$currentSheet->getCell('V'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>内在动机</td> <td>22</td> <td>24</td> <td>25</td> <td>'.$currentSheet->getCell('W'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>外在动机</td> <td>44.25</td> <td>50</td> <td>55</td> <td>'.$currentSheet->getCell('X'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>团队倾向</td> <td>11</td> <td>12</td> <td>13</td> <td>'.$currentSheet->getCell('Y'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>创新效能感</td> <td>19</td> <td>21</td> <td>23</td> <td>'.$currentSheet->getCell('Z'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>团队中自尊感</td> <td>50</td> <td>55</td> <td>60.75</td> <td>'.$currentSheet->getCell('AA'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>风险规避倾向</td> <td>28</td> <td>32</td> <td>37</td> <td>'.$currentSheet->getCell('AB'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>开放性</td> <td>45.25</td> <td>49</td> <td>53</td> <td>'.$currentSheet->getCell('AC'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>尽责性</td> <td>40</td> <td>44</td> <td>47</td> <td>'.$currentSheet->getCell('AD'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>掌握目标倾向</td> <td>26</td> <td>29</td> <td>31</td> <td>'.$currentSheet->getCell('AE'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>表现目标倾向</td> <td>20</td> <td>23</td> <td>26.75</td> <td>'.$currentSheet->getCell('AF'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>回避目标倾向</td> <td>12</td> <td>15</td> <td>17</td> <td>'.$currentSheet->getCell('AG'.$row)->getValue().'</td> </tr>';
			//下面这两个数据有问题
			//$mailcon=$mailcon.'<tr> <td>多任务</td> <td>44</td> <td>47</td> <td>53</td> <td>'.$currentSheet->getCell('AH'.$row)->getValue().'</td> </tr>';
			//$mailcon=$mailcon.'<tr> <td>不确定规避</td> <td>25</td> <td>28</td> <td>30</td> <td>'.$currentSheet->getCell('AI'.$row)->getValue().'</td> </tr>';
			
			$mailcon=$mailcon.'<tr> <td>个人-偏好集体</td> <td>32</td> <td>35</td> <td>37</td> <td>'.$currentSheet->getCell('AH'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>权力距离意识</td> <td>11</td> <td>14</td> <td>18</td> <td>'.$currentSheet->getCell('AI'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>观点采择能力</td> <td>20</td> <td>23</td> <td>24</td> <td>'.$currentSheet->getCell('AJ'.$row)->getValue().'</td> </tr>';
			$mailcon=$mailcon.'<tr> <td>创新卷入行为</td> <td>53</td> <td>58</td> <td>63</td> <td>'.$currentSheet->getCell('AK'.$row)->getValue().'</td> </tr>';

			$mailcon=$mailcon.'</table>';

			$mailcon=$mailcon.'<h2>名词解释</h2><p>突破创新倾向：希望有突破性创新表现的倾向。</p><p>责任意识：对集体、对他人表现出来的责任心。</p><p>服从倾向：对权威或其他重要人物表现出来的服从性。</p><p>渐进创新倾向：对一般、小改进的创新偏好的倾向。</p><p>积极进取意识：积极进取性。</p><p>内在动机：对任务本身的兴趣；</p><p>外在动机：对因任务带来的金钱、荣誉或其他外部奖励的兴趣。</p><p>团队倾向：你对在团队中工作的偏好程度。</p><p>创新效能感：对自己的创新能力的自信程度。</p><p>团队中的自尊感：在团队中你对自己的影响力和重要性的判断。</p><p>风险规避倾向：你在工作或学习中表现出的对风险的回避性。在生活的选择中也会有这样的倾向。</p><p>开放性：具有想象、审美、情感丰富、求异、创造、智慧等特征。得分越高，你的这些特征表现越明显。</p><p>尽责性：胜任、公正、条理、尽职、成就、自律、谨慎、克制等特点。得分越高，你的这些特征表现越明显。</p><p>掌握目标倾向：在进行工作或学习时，希望真正掌握知识或技能的倾向。</p><p>表现目标倾向：在进行工作或学习时，为了表现自己的能力的倾向。</p><p>回避目标倾向：在进行工作或学习时，尽力回避失败的倾向。</p><p>个人-集体偏好：你是偏好个人工作还是团队工作。得分越高，表明越偏向集体工作。</p><p>权力距离意识：对组织中权力分配不平等情况的接受程度。权力距离越大，表明你可以接受领导有特权的容忍度越大。</p><p>观点采择能力：能站在他人角度思考问题的能力。</p><p>创新卷入行为：你在创新团队实际表现出的创新行为。</p>';

			//发送邮件
			yqc_mail($mailto,$mailsub,$mailcon);
			sleep(9);
		}
			

		?>		
	</div>
</div>
</body>
</html>

