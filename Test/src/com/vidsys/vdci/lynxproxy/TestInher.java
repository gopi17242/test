package com.vidsys.vdci.lynxproxy;
public class TestInher {

	public static void main(String[] args) {
		Super s = new Subclass();
		s.foo();
	}
}

class Super {
	void foo() {
		System.out.println("Super");
	}
}

class Subclass extends Super {
	static void foo() {
		System.out.println("Subclass");
	}

}