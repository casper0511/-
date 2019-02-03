package eng;

import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

public class test1 {


	public static void main(String[] args) {
		List<String> list = new LinkedList<String>();
		for (int i = 0; i < 9; i++) {
			list.add("a" + i);
		}

		Collections.shuffle(list); // »ìÂÒµÄÒâË¼
		System.out.println(list);

	}

}