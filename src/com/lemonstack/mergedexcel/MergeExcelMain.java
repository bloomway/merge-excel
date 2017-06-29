package com.lemonstack.mergedexcel;

import java.io.File;
import java.util.Arrays;
import java.util.List;

public class MergeExcelMain {

	public static void main(String[] args) {
		
		
		final MergeManager merger = new MergeManager();
		
		final File src = new File(args[0]);
		final File dst = new File(args[1]);
		
		final List<File> files = Arrays.asList(src.listFiles());
		System.out.println("Start merging...");
		merger.merge(files, dst);
		
		System.out.println("...Done!");
	}

}
