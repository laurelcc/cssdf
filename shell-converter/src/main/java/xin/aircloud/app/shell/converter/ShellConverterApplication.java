package xin.aircloud.app.shell.converter;

import org.jline.utils.AttributedString;
import org.jline.utils.AttributedStyle;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.shell.jline.PromptProvider;

@SpringBootApplication
public class ShellConverterApplication {

	public static void main(String[] args) {
		SpringApplication.run(ShellConverterApplication.class, args);
	}

    /**
     * Prompt提示符
     * @return
     */
	@Bean
    public PromptProvider promptProvider(){
	    return () -> new AttributedString("soong-converter:>",
                AttributedStyle.DEFAULT.foreground(AttributedStyle.YELLOW));
    }

}
