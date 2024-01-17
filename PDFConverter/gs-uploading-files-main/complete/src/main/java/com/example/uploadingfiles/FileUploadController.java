package com.example.uploadingfiles;

import java.io.File;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.tomcat.util.http.fileupload.FileUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.web.ServerProperties.Undertow.Options;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.MvcUriComponentsBuilder;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.aspose.slides.Presentation; //PPT-TO-PDF ----- WORKING FINE
import com.aspose.slides.*;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
//import com.aspose.slides.*; //PPT-TO-PDF ----- WORKING FINE
import com.example.uploadingfiles.storage.StorageFileNotFoundException;
import com.example.uploadingfiles.storage.StorageProperties;
import com.example.uploadingfiles.storage.StorageService;
import com.spire.doc.Document; //DOC-TO-PDF ----- WORKING FINE
import com.spire.doc.FileFormat; //DOC-TO-PDF ------ WORKING FINE
import com.spire.pdf.PdfDocument;
import com.spire.pdf.PdfDocumentBase;
import com.spire.pdf.*;


@Controller
public class FileUploadController {

	private final StorageService storageService;

	@Autowired
	public FileUploadController(StorageService storageService) {
		this.storageService = storageService;
	}

	@GetMapping("/home")
	public String indexPage() {
		return "home";
	}

	@GetMapping("/about")
	public String contactPage() {
		return "about";
	}

	@GetMapping("/formpage")
	public String formPage() {
		return "formpage";
	}

	@GetMapping("/doc-to-pdf")
	public String DocToPdf(Model model) throws IOException {
		
		model.addAttribute("welcome", "Convert Documents to PDF");
		model.addAttribute("files", storageService.loadAll().map(
				path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
						"serveFile", path.getFileName().toString()).build().toUri().toString())
				.collect(Collectors.toList()));

		return "doc-to-pdf";
	}

	@GetMapping("/ppt-to-pdf")
	public String PptToPdf(Model model) throws IOException {
		model.addAttribute("welcome", "Convert Presentations to PDF");
		model.addAttribute("files", storageService.loadAll().map(
				path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
						"serveFile", path.getFileName().toString()).build().toUri().toString())
				.collect(Collectors.toList()));

		return "ppt-to-pdf";
	}

	@GetMapping("/exc-to-pdf")
	public String ExcToPdf(Model model) throws IOException {
		model.addAttribute("welcome", "Convert Excel to PDF");
		model.addAttribute("files", storageService.loadAll().map(
				path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
						"serveFile", path.getFileName().toString()).build().toUri().toString())
				.collect(Collectors.toList()));

		return "exc-to-pdf";
	}

	@GetMapping("/merge-pdfs")
	public String mergePdf(Model model) throws IOException {
		model.addAttribute("welcome", "Merge Multiple PDFs");
		model.addAttribute("files", storageService.loadAll().map(
				path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
						"serveFile", path.getFileName().toString()).build().toUri().toString())
				.collect(Collectors.toList()));

		return "merge-pdf";
	}

	public String listUploadedFiles(Model model) throws IOException {

		model.addAttribute("files", storageService.loadAll().map(
				path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
						"serveFile", path.getFileName().toString()).build().toUri().toString())
				.collect(Collectors.toList()));

		return "uploadForm";
	}

	@GetMapping("/files/{filename:.+}")
	@ResponseBody
	public ResponseEntity<Resource> serveFile(@PathVariable String filename) {

		Resource file = storageService.loadAsResource(filename);
		return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION,
				"attachment; filename=\"" + file.getFilename() + "\"").body(file);
	}

	@PostMapping("/doc-to-pdf")
	public String handleDocUpload(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) {
		storageService.store(file);
		convertDocToPdf(file.getOriginalFilename());
		redirectAttributes.addFlashAttribute("message",
				"You successfully uploaded " + file.getOriginalFilename() + "!");

		return "redirect:/doc-to-pdf";
	}	

	@PostMapping("/ppt-to-pdf")
	public String handlePptUpload(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) {
		storageService.store(file);
		convertPptToPdf(file.getOriginalFilename());
		redirectAttributes.addFlashAttribute("message",
				"You successfully uploaded " + file.getOriginalFilename() + "!");

		return "redirect:/doc-to-pdf";
	}

	@PostMapping("/exc-to-pdf")
	public String handleExcUpload(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) throws Exception {
		storageService.store(file);
		convertExcToPdf(file.getOriginalFilename());
		redirectAttributes.addFlashAttribute("message",
				"You successfully uploaded " + file.getOriginalFilename() + "!");
		return "redirect:/exc-to-pdf";
	}

	@PostMapping("/merge-pdf")
	public String handleMergeUpload(@RequestParam("file") MultipartFile[] files, RedirectAttributes redirectAttributes) {
		List<String> fileNames = new ArrayList<>();
		Arrays.asList(files).stream().forEach(file -> {
			storageService.store(file);
			fileNames.add(file.getOriginalFilename());
		});
		mergePdf(fileNames);
		redirectAttributes.addFlashAttribute("message",
				"You successfully uploaded!");
		return "redirect:/merge-pdfs";
	}

	public void convertDocToPdf(String fileName){
		Document doc = new Document("upload-dir\\" + fileName);
		doc.saveToFile("upload-dir\\" + fileName + "(output).pdf", FileFormat.PDF);
	}
	public void convertPptToPdf(String fileName){
		Presentation presentation = new Presentation("upload-dir\\" + fileName);
		presentation.save("upload-dir\\" + fileName + "(output).pdf", SaveFormat.Pdf);
	}
	public void convertExcToPdf(String fileName) throws Exception{
		Workbook workbook = new Workbook("upload-dir\\" + fileName);
		PdfSaveOptions options = new PdfSaveOptions();
		options.setCompliance(PdfCompliance.PdfA1a);
		workbook.save("upload-dir\\" + fileName + "(output).pdf", options);
	}
	public void mergePdf(List<String> files) {
		String s[] = new String[files.size()];
		for(int i=0; i<files.size(); i++) {
			s[i] = "upload-dir\\" + files.get(i);
		}
		PdfDocumentBase doc = PdfDocument.mergeFiles(s);
		doc.save("upload-dir\\output.pdf");
	}

	@ExceptionHandler(StorageFileNotFoundException.class)
	public ResponseEntity<?> handleStorageFileNotFound(StorageFileNotFoundException exc) {
		return ResponseEntity.notFound().build();
	}

}
