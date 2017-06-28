package com.lemonstack.mergedexcel;

import java.io.File;

public class Main {

	public static void main(String[] args) {
		
		System.out.println("Start merging...");
		final Merger merger = new Merger();
		
		merger.merge(new File(args[0]));
		
		System.out.println("...Done!");
	}

}
