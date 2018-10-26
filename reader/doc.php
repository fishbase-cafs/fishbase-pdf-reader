<?php
function curl($url){
        $ch = curl_init();
        curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, 5000);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($ch, CURLOPT_HTTPHEADER, array('User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0'));
        curl_setopt($ch, CURLOPT_URL, $url);
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
        curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
        $contents = curl_exec($ch);
        curl_close($ch);//关闭一打开的会话
        return $contents;
}
function getAPI_CDN($url_encoded)
{
	$routine=curl("https://view.officeapps.live.com/op/view.aspx?src=".$url_encoded);
	$urlro=strstr(str_replace('\u0025','%',strstr($routine,'WOPISrc=')),'wFileId%3Dhttp',true)."wFileId%3D";
	return $urlro;
}
function getWordPDFUrl($docurl)
{
	$url=urlencode($docurl);
	$urlro=getAPI_CDN($url);
	$urlwo='https://word-view.officeapps.live.com/wv/wordviewerframe.aspx?'.$urlro.urlencode($url).'&access_token_ttl=0';
	$routpdf=curl($urlwo);
	$coder='https://word-view.officeapps.live.com/wv/WordViewer/' . str_replace('ResReader.ashx','Document.pdf',strstr(strstr($routpdf,'ResReader.ashx'),'"',true))."&wdAccPdf=1&access_token=1&access_token_ttl=0&type=downloadpdf";
	return $coder;
}

$http_type = ((isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] == 'on') || (isset($_SERVER['HTTP_X_FORWARDED_PROTO']) && $_SERVER['HTTP_X_FORWARDED_PROTO'] == 'https')) ? 'https://' : 'http://';

$path=trim($_SERVER["QUERY_STRING"]);
if(strlen($path)==0)die();
$docpath=substr($_SERVER['PHP_SELF'],0,strripos($_SERVER['PHP_SELF'],"/")+1);
$upath=$docpath.$path;
if($path[0]=='/')
{
	$upath=$path;
};
echo $upath;

$url=$http_type . $_SERVER['HTTP_HOST'].'/'.$upath;
if(strpos(strtolower($url),'.doc'))
{
$tmpfile=base64_encode($url);

if (function_exists('apcu_exists')) {
		if(apcu_exists($tmpfile))
		{
		       $data=apcu_fetch($tmpfile);
		       if(strpos($data,'PDF')==1)
			{
				header('Content-type: application/pdf');
				echo $data;
			}
		}else
		{
		       $data=curl(getWordPDFUrl($url));
			if(strpos($data,'PDF')==1)
	                {
			       header('Content-type: application/pdf');
			       apcu_add($tmpfile,$data,3600);
			       echo $data;
			}
		}
	} else {
	       $data=curl(getWordPDFUrl($url));
	       header('Content-type: application/pdf');
	       echo $data;
	}
}
?>
