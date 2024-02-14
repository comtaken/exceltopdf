package com.example.demo.controller;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.util.UriUtils;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.example.demo.form.InputForm;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
@Controller
public class OutptExcelToPdf {
	@Autowired
	private MessageSource messageSource;
	/**
	 * フォーム画面表示
	 * @param model
	 * @return input-form
	 * @throws IOException 
	 */
	@GetMapping("/")
	public String displayForm(Model model){
		String confirmMessage = messageSource.getMessage("confirm.message", null, Locale.JAPAN);
		model.addAttribute("inputform", new InputForm());
		model.addAttribute("confirmMessage",confirmMessage );
		return "input-form";
	}

	/**
	 * PDF出力処理
	 * @param inputform
	 * @param model
	 * @return フォーム内容の印字されたPDF
	 * @throws Exception
	 */
	@PostMapping("/export")
	public String exportForm(@ModelAttribute InputForm inputform, Model model, RedirectAttributes redirectAttributes,HttpServletRequest request,HttpServletResponse response)
			throws Exception {
		FileOutputStream contents = null;
		XSSFWorkbook workbook = null;
		System.out.println("excelへ出力開始");
		String filePath = "";
		//ファイル原本コピー
		Path originalPath = Paths.get("./form.xlsx");
		Path copyPath = Paths.get("./form_tmp.xlsx");
		try {
			Files.copy(originalPath, copyPath);
		} catch (IOException e) {
			System.out.println(e);
		}
		try {
			filePath = copyPath.toString();
			FileInputStream fis = new FileInputStream(filePath);
			workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			//テンプレート内を走査して置換を行う
			for (Row rowi : sheet) {
				for (Cell cellj : rowi) {
					//日程
					if ("#SCHEDULE".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getSCHEDULE());
					}
					//出発地点
					if ("#STARTPOINT".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getSTARTPOINT());
					}
					//到着地点
					if ("#ENDPOINT".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getENDPOINT());
					}
					//経由１
					if ("#VIA1".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getVIA1());
					}
					//経由２
					if ("#VIA2".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getVIA2());
					}
					//経由３
					if ("#VIA3".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getVIA3());
					}
					//経由４
					if ("#VIA4".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getVIA4());
					}
					//経由５
					if ("#VIA5".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getVIA5());
					}
					//経由６
					if ("#VIA6".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getVIA6());
					}
					//食事：朝
					if ("#BREAKFAST".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getBREAKFAST());
					}
					//食事：昼
					if ("#LUNCH".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getLUNCH());
					}
					//食事：夜
					if ("#DINNER".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getDINNER());
					}
					//宿泊場：施設名
					if ("#HOTEL".equals(cellj.getStringCellValue())) {
						cellj.setCellValue(inputform.getHOTEL());
					}
				}
			}
			contents = new FileOutputStream(filePath);
			workbook.write(contents);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (contents != null) {
					contents.close();
				}
				if (workbook != null) {
					workbook.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		Path path = Paths.get("src/main/resources/application.properties");
		Reader reader = Files.newBufferedReader(path, StandardCharsets.UTF_8);
		Properties properties = new Properties();
		properties.load(reader);
		String pdfFolderPath = properties.getProperty("pdfFolderPath");
		//form_tmp.xlsxをPDF出力
		Workbook workbooktopdf = new Workbook(filePath);
		LocalDateTime nowDate = LocalDateTime.now();
		DateTimeFormatter formater = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
		String yyyyMMddHHmmss = formater.format(nowDate);
		
		workbooktopdf.save(pdfFolderPath + "ToPDF" + yyyyMMddHHmmss + ".pdf", SaveFormat.PDF);
		System.out.println("export-pdf\\ToPDF" + yyyyMMddHHmmss + ".pdfを出力しました。");
		//PDF出力後はform_tmp.xlsxを削除
		try {
			Files.delete(copyPath);
			String message = "出力完了しました。";
			model.addAttribute("inputform", new InputForm());
			redirectAttributes.addFlashAttribute("message", message);
		} catch (IOException e) {
			System.out.println(e);
		}
		
		downloadPdf(pdfFolderPath + "ToPDF" + yyyyMMddHHmmss + ".pdf","ToPDF" + yyyyMMddHHmmss + ".pdf", response);
		System.out.println("ダウンロード完了");
		return null;
	}
	/**
	 * pdfダウンロード
	 * @param originFilePath
	 * @param outputFileName
	 * @param response
	 */
	public static void downloadPdf(String originFilePath, String outputFileName, HttpServletResponse response) {
		String contentFormat = "attachment; filename=\"%s\"; filename*=UTF-8''%s";
		outputFileName = String.format(contentFormat, outputFileName,
				UriUtils.encode(outputFileName, StandardCharsets.UTF_8.name()));
		try (OutputStream os = response.getOutputStream();) {
			Path filePath = Path.of(originFilePath);
			byte[] biteFile = Files.readAllBytes(filePath);
			response.setHeader(HttpHeaders.CONTENT_DISPOSITION, outputFileName);
			response.setContentType("application/octet-stream");
			response.setContentLength(biteFile.length);
			os.write(biteFile);
			os.flush();
			os.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
