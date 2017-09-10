package com.test.KeywordExtractor;

import java.io.IOException;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.lucene.analysis.ASCIIFoldingFilter;
import org.apache.lucene.analysis.LowerCaseFilter;
import org.apache.lucene.analysis.PorterStemFilter;
import org.apache.lucene.analysis.StopFilter;
import org.apache.lucene.analysis.TokenStream;
import org.apache.lucene.analysis.en.EnglishAnalyzer;
import org.apache.lucene.analysis.standard.ClassicFilter;
import org.apache.lucene.analysis.standard.ClassicTokenizer;
import org.apache.lucene.analysis.tokenattributes.CharTermAttribute;
import org.apache.lucene.util.Version;

public class KeywordExtractor {
	
	public static List<Keyword> keywords;

	public String stem(String term) throws IOException {
		TokenStream tokenStream = null;

		try {
			tokenStream = new ClassicTokenizer(Version.LUCENE_36, new StringReader(term));
			tokenStream = new PorterStemFilter(tokenStream);

			Set<String> stems = new HashSet<String>();
			CharTermAttribute token = tokenStream.getAttribute(CharTermAttribute.class);
			tokenStream.reset();

			while (tokenStream.incrementToken()) {
				stems.add(token.toString().toString());
			}

			if (stems.size() != 1) {
				return null;
			}
			String stem = stems.iterator().next();

			if (!stem.matches("[a-zA-Z0-9-]+")) {
				return null;
			}

			return stem;
		} finally {
			if (tokenStream != null) {
				tokenStream.close();
			}
		}
	}

	public <T> T find(Collection<T> collection, T example) {
		for (T element : collection) {
			if (element.equals(example)) {
				return element;
			}
		}
		collection.add(example);
		return example;
	}
	
	public String cleanInput(String input) {
		input = input.replaceAll("-+", "-0");
		input = input.replaceAll("[\\p{Punct}&&[^'-]]+", " ");
		input = input.replaceAll("(?:'(?:[tdsm]|[vr]e|ll))+\\b","");		
		return input;
	}
	
	public TokenStream tokenizeStream(String input) {
		return new ClassicTokenizer(Version.LUCENE_36, new StringReader(input));
	}
	
	public TokenStream filterStream(TokenStream tokenStream) {
		TokenStream result = new LowerCaseFilter(Version.LUCENE_36, tokenStream);
		result = new ClassicFilter(tokenStream);
		tokenStream = new LowerCaseFilter(Version.LUCENE_36, tokenStream);
		result = new ASCIIFoldingFilter(tokenStream);
		result = new StopFilter(Version.LUCENE_36, tokenStream, EnglishAnalyzer.getDefaultStopSet());
		return result;
	}
	
	public void addKeywords(String term, int weight) throws IOException {		
		String stem = stem(term);
		if (stem != null && stem.length() >= 2 && stem.matches("[a-zA-Z]+")) {
			Keyword keyword = find(keywords, new Keyword(stem.replaceAll("-0", "-")));
			keyword.add(term.replaceAll("-0", "-"), weight);
		}
	}
	
	public void processAndAddText(String input, int weight) throws IOException {
		String cleanedInput = cleanInput(input);
		TokenStream tokenStream = tokenizeStream(cleanedInput);
		tokenStream = filterStream(tokenStream);
		CharTermAttribute token = tokenStream.getAttribute(CharTermAttribute.class);
		tokenStream.reset();
		while(tokenStream.incrementToken()) {
			addKeywords(token.toString(), weight);
		}
	}
	
	public void processSlide(Slide slide) throws IOException {
		processAndAddText(slide.getText(), 1);
		processAndAddText(slide.getBoldText(), 1);
		processAndAddText(slide.getTitleText(), 3);
	}
	
	public void processPresentation(Presentation presentation) throws IOException {
		for (Slide slide : presentation.getSlides()) {
			processSlide(slide);
		}
	}
	
	public void sortKeywords(){
		Collections.sort(keywords);
	}
	
	public KeywordExtractor() {
		keywords = new ArrayList<Keyword>();
	}

	public static void main(String[] args) throws IOException {
		String filename = args[0];
		int k = Integer.parseInt(args[1]);
		
		PPTParser pptParser = new PPTParser();
		Presentation presentation = pptParser.parsePPT(filename);
		
		KeywordExtractor keywordExtractor = new KeywordExtractor();
		keywordExtractor.processPresentation(presentation);
		
		keywordExtractor.sortKeywords();
		
		for (Keyword keyword:keywords) {
			if (k < 0) break;
			System.out.println(keyword.getStem() + "\t" + keyword.getFrequency() + "\t" + keyword.getTerms());
			k--;
		}
	}

}
