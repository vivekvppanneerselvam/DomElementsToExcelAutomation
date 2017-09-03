package com.dom.byoextractor;

import java.io.IOException;

import org.jsoup.Connection;
import org.jsoup.Jsoup;

public class BYOExtractorConfig {
	
	public static Connection.Response confluenceConfig () throws IOException{
		System.setProperty("http.proxyHost", "please provide proxy");
	    System.setProperty("http.proxyPort", "please provide port");
	    Connection.Response res = Jsoup
	     	    .connect("provide url")
	     	    // if credential required 
	     	    .data("username field id", "provide input here")
	     	    .data("password field id", "provide input here")
	     	    .data("button id", "give button value here")
	     	    .method(Connection.Method.POST).followRedirects(false)
	     	    .execute();
		return res;
	}

}


