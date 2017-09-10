package com.test.KeywordExtractor;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

class Slide {
	String text;
	String boldText;
	String titleText;
	
	public Slide(String text, String boldText, String titleText) {
		this.text = text;
		this.boldText = boldText;
		this.titleText = titleText;
	}
	
	public void setText(String text) {
		this.text = text;
	}
	
	public void setBoldText(String text) {
		this.boldText = text;
	}
	
	public void setTitleText(String text) {
		this.titleText = text;
	}
	
	public String getText() {
		return this.text;
	}
	
	public String getBoldText() {
		return this.boldText;
	}
	
	public String getTitleText() {
		return this.titleText;
	}
	
}

class Presentation {
	ArrayList<Slide> slides;
	
	public Presentation() {
		slides = new ArrayList<Slide>();
	}
	
	public void addSlide(String text, String boldText, String titleText) {
		if (text == null) text = "";
		if (boldText == null) boldText = "";
		if (titleText == null) titleText = "";
		Slide slide = new Slide(text, boldText, titleText);
		slides.add(slide);
	}
	
	public void addSlide(XSLFSlide slide) {
		String titleText = slide.getTitle();
		String boldText = "";
		String text = "";
		for(XSLFShape shape : slide.getShapes()){
			if (shape instanceof XSLFTextShape) {
				XSLFTextShape tsh = (XSLFTextShape)shape;
				text += tsh.getText() + " ";
				for (XSLFTextParagraph p : tsh) {
					for (XSLFTextRun r : p) {
						if (r.isBold()) {
							boldText += r.getRawText() + " ";
						}
					}
				}
			}
		}
		addSlide(text, boldText, titleText);
	}
	
	public ArrayList<Slide> getSlides() {
		return this.slides;
	}

}

public class PPTParser {
	
	private static XMLSlideShow ppt;
	
	public Presentation parsePPT(String filename) {
		Presentation presentation = new Presentation();
		try {
			FileInputStream fis = new FileInputStream(filename);
			ppt = new XMLSlideShow(fis);
			fis.close();			
			for (XSLFSlide slide : ppt.getSlides()) {
				presentation.addSlide(slide);
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return presentation;
	}
	
}
