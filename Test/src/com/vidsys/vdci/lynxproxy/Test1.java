package com.vidsys.vdci.lynxproxy;

import java.nio.file.Path;
import java.nio.file.Paths;

public class Test1 {

	public static void main(String[] args) {

//		final String dir = System.getProperty("user.dir");
		Path workingDirectory=Paths.get(".").toAbsolutePath();
        System.out.println("current dir = " + workingDirectory.getRoot());
        System.out.println("current dir = " + workingDirectory);
	}

}
