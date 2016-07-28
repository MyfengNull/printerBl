package com.example.printer;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager.NameNotFoundException;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.Toast;

/**
 * @time 2015年12月14日08:28:28
 * @author osy
 * @version 1.0
 */
public class PrinterActivity extends Activity implements OnClickListener {
	private Context context;
	private Button prinrerBtn, btnAz;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.printer);
		prinrerBtn = (Button) findViewById(R.id.button1);
		btnAz = (Button) findViewById(R.id.btnAz);
		prinrerBtn.setOnClickListener(this);
		btnAz.setOnClickListener(this);

	}

	// 打印检查记录表
	/**
	 * 为了保证模板的可用，最好在现有的模板上复制后修改
	 */
	private void printer() {
		try {
			saveFile("xcjcjl.doc", PrinterActivity.this, R.raw.xcjcjl);// 文件目录res/raw
		} catch (IOException e) {

			e.printStackTrace();
		}
		// 现场检查记录
		String aafileurl = Environment.getExternalStorageDirectory()
				+ "/inspection/xcjcjl.doc";
		final String bbfileurl = Environment.getExternalStorageDirectory()
				+ "/inspection/xcjcjl_printer.doc";
		// 获取模板文件
		File demoFile = new File(aafileurl);
		// 创建生成的文件
		File newFile = new File(bbfileurl);
		if (newFile.exists()) {
			newFile.delete();
		}
		Map<String, String> map = new HashMap<String, String>();
		map.put("$record_companyName$", "涉及项目不提供");
		map.put("$record_companyAddress$", "涉及项目不提供");
		map.put("$record_companyPic$", "涉及项目不提供");
		map.put("$record_companyWork$", "涉及项目不提供");
		map.put("$record_companyPhone$", "涉及项目不提供");
		map.put("$record_CheckAddress$", "涉及项目不提供");
		map.put("$time_nian$", "涉及项目不提供");
		map.put("$time_yue$", "涉及项目不提供");
		map.put("$time_ri$", "涉及项目不提供");
		map.put("$time_shi$", "涉及项目不提供");
		map.put("$time_fen$", "涉及项目不提供");
		map.put("$time_ri2$", "涉及项目不提供");
		map.put("$time_shi2$", "涉及项目不提供");
		map.put("$time_fen2$", "涉及项目不提供");
		map.put("$record_jcjg$", "涉及项目不提供");
		map.put("$record_userName$", "涉及项目不提供");
		map.put("$record_userName2$", "涉及项目不提供");
		map.put("$record_userNum$", "涉及项目不提供");
		map.put("$record_userNum2$", "涉及项目不提供");
		map.put("$content$", "涉及项目不提供");
		if (writeDoc(demoFile, newFile, map)) {

			// 调用printershare软件来打印该文件
			File picture = new File(bbfileurl);
			Uri data_uri = Uri.fromFile(picture);
			/*
			 * data_type - Mime type. MIME类型如下:
			 * 
			 * "application/pdf" "text/html" "text/plain" "image/png"
			 * "image/jpeg" "application/msword" - .doc
			 * "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
			 * - .docx "application/vnd.ms-excel" - .xls
			 * "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
			 * - .xlsx "application/vnd.ms-powerpoint" - .ppt
			 * "application/vnd.openxmlformats-officedocument.presentationml.presentation"
			 * - .pptx
			 */
			try {

				String data_type = "application/msword";
				Intent i = new Intent(Intent.ACTION_VIEW);
				i.setPackage("com.dynamixsoftware.printershare");// 未注册之前com.dynamixsoftware.printershare，注册后加上amazon
				i.setDataAndType(data_uri, data_type);
				startActivity(i);
			} catch (Exception e) {
				// 没有找到printershare
				Log.e("TAG", "没有找到printershare");
			}
		}

	}

	/**
	 * demoFile 模板文件 newFile 生成文件 map 要填充的数据
	 * */
	public boolean writeDoc(File demoFile, File newFile, Map<String, String> map) {
		try {
			FileInputStream in = new FileInputStream(demoFile);
			HWPFDocument hdt = new HWPFDocument(in);
			// Fields fields = hdt.getFields();
			// 读取word文本内容
			Range range = hdt.getRange();
			// System.out.println(range.text());

			// 替换文本内容
			for (Map.Entry<String, String> entry : map.entrySet()) {
				range.replaceText(entry.getKey(), entry.getValue());
			}
			ByteArrayOutputStream ostream = new ByteArrayOutputStream();
			FileOutputStream out = new FileOutputStream(newFile, true);
			hdt.write(ostream);
			// 输出字节流
			out.write(ostream.toByteArray());
			out.close();
			ostream.close();
			return true;
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return false;
	}

	/**
	 * 将文件复制到SD卡，并返回该文件对应的数据库对象
	 * 
	 * @return
	 * @throws IOException
	 */
	public void saveFile(String fileName, Context context, int rawid)
			throws IOException {

		// 首先判断该目录下的文件夹是否存在
		File dir = new File(Environment.getExternalStorageDirectory()
				+ "/inspection/");
		if (!dir.exists()) {
			// 文件夹不存在 ， 则创建文件夹
			dir.mkdirs();
			Toast.makeText(getApplication(), "文件夹不存在", Toast.LENGTH_SHORT)
					.show();
		}

		// 判断目标文件是否存在
		File file1 = new File(dir, fileName);

		if (!file1.exists()) {
			file1.createNewFile(); // 创建文件
			Toast.makeText(getApplication(), "创建文件", Toast.LENGTH_SHORT).show();
		}
		// 开始进行文件的复制
		InputStream input = context.getResources().openRawResource(rawid); // 获取资源文件raw
																			// 标号
		try {

			FileOutputStream out = new FileOutputStream(file1); // 文件输出流、用于将文件写到SD卡中
																// -- 从内存出去
			byte[] buffer = new byte[1024];
			int len = 0;
			while ((len = (input.read(buffer))) != -1) { // 读取文件，-- 进到内存

				out.write(buffer, 0, len); // 写入数据 ，-- 从内存出
			}

			input.close();
			out.close(); // 关闭流
		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	// 判断apk是否安装
	public static boolean appIsInstalled(Context context, String pageName) {
		try {
			context.getPackageManager().getPackageInfo(pageName, 0);
			return true;
		} catch (NameNotFoundException e) {
			return false;
		}
	}

	// 把Asset下的apk拷贝到sdcard下 /Android/data/你的包名/cache 目录下
	public static File getAssetFileToCacheDir(Context context, String fileName) {
		try {
			File cacheDir = FileUtils.getCacheDir(context);
			final String cachePath = cacheDir.getAbsolutePath()
					+ File.separator + fileName;
			InputStream is = context.getAssets().open(fileName);
			File file = new File(cachePath);
			file.createNewFile();
			FileOutputStream fos = new FileOutputStream(file);
			byte[] temp = new byte[1024];

			int i = 0;
			while ((i = is.read(temp)) > 0) {
				fos.write(temp, 0, i);
			}
			fos.close();
			is.close();
			return file;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	// 获取sdcard中的缓存目录
	public static File getCacheDir(Context context) {
		String APP_DIR_NAME = Environment.getExternalStorageDirectory()
				.getAbsolutePath() + "/Android/data/";
		File dir = new File(APP_DIR_NAME + context.getPackageName() + "/cache/");
		if (!dir.exists()) {
			dir.mkdirs();
		}
		return dir;
	}

	public boolean copyApkFromAssets(Context context, String fileName,
			String path) {
		boolean copyIsFinish = false;
		try {
			InputStream is = context.getAssets().open("PrinterShare-11.0.0.apk");
			File file = new File(path);
			file.createNewFile();
			FileOutputStream fos = new FileOutputStream(file);
			byte[] temp = new byte[1024];
			int i = 0;
			while ((i = is.read(temp)) > 0) {
				fos.write(temp, 0, i);
			}
			fos.close();
			is.close();
			copyIsFinish = true;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return copyIsFinish;
	}

	@Override
	public void onClick(View v) {
		// TODO Auto-generated method stub
		switch (v.getId()) {
		case R.id.button1:
			printer();
			break;
		case R.id.btnAz:
			if(copyApkFromAssets(this, "PrinterShare-11.0.0.apk", Environment.getExternalStorageDirectory().getAbsolutePath()+"/PrinterShare-11.0.0.apk")){
				Intent intent = new Intent(Intent.ACTION_VIEW);
				intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
			    String	path=Environment.getExternalStorageDirectory().getAbsolutePath();
				intent.setDataAndType(
						Uri.parse("file://"+ Environment.getExternalStorageDirectory()
								.getAbsolutePath() + "/PrinterShare-11.0.0.apk"),
						"application/vnd.android.package-archive");
				Log.e("path", path);
				startActivity(intent);
			}
			break;
		default:
			break;
		}
	}
	 // 判断SDCard是否存在
    private boolean isHaveSDCard() {
        String status = Environment.getExternalStorageState();
        if (status.equals(Environment.MEDIA_MOUNTED)) {
            return true;
        } else {
            return false;
        }
    }
}
