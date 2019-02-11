package eng;

import java.io.File;

public class Wordlist {
	
	private String word;
	private File path;
	
	public Wordlist(String word, File path) {
        this.word = word;
        this.path = path;
    }
	
	public String getWord() {
        return word;
    }
    public void setWord(String word) {
        this.word = word;
    }
    public File getPath() {
        return path;
    }
    public void setPath(File path) {
        this.path = path;
    }


}
 