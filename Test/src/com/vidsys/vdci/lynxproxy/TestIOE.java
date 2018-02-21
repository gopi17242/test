package com.vidsys.vdci.lynxproxy;

import java.io.IOException;

public class TestIOE {

	public static void main(String[] args) throws IOException {
		try {
		//	throw new IOException("Hello");
		}catch(Exception e) {
			System.out.println(e.getMessage());
		}
	}
}