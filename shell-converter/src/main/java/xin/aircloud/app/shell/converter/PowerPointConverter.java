package xin.aircloud.app.shell.converter;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;
import org.springframework.shell.standard.ShellOption;
import org.springframework.util.StringUtils;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicInteger;

@ShellComponent(value = "Convert the snapshot of pptx/ppt to png. Type help for more infomation.")
public class PowerPointConverter {

    private static Logger logger = LoggerFactory.getLogger(PowerPointConverter.class);

    @ShellMethod(value = "snapshots of ppt/pptx to png", key = "ppt2png", group = "PPTx to PNG Converter")
    public void convertSingleFile(@ShellOption(value = "file", defaultValue = "") String file,
                                  @ShellOption(value = "cols", defaultValue = "3") int cols,
                                  @ShellOption(value = "space", defaultValue = "0") int space,
                                  @ShellOption(value = "water", defaultValue = "") String warter) throws Exception{
        if (StringUtils.isEmpty(file)){
            print("Please confirm the target file extension pptx or ppt");
        } else {
            Path filePath = Paths.get(file);
            File targetFile = new File(filePath.toUri());
            if (!targetFile.exists()){
                print("The PPTx file dose not exists");
            } else {
                if (!targetFile.isFile()){
                    print("Not a file type");
                } else {
                    if (!checkPPTxExtension(targetFile.getName())){
                        print("Not a PPTx file");
                    }
                }
            }
        }

        int colsCount = cols;
        if (cols < 0 || cols > 12) {
            print("The cols should be in 1...12");
        }

        int spaceNum = space;
        if (space < 0 || space > 60) {
            print("The space between cols should be in 0...60");
        }

        if (!StringUtils.isEmpty(warter)){
            Path filePath = Paths.get(warter);
            File targetFile = new File(filePath.toUri());
            if (!targetFile.exists()){
                print("The wartermark file dose not exists");
            }
        }

        print("\nprocessing...");

        process(file, colsCount, spaceNum);

        print("\nJob Done");
    }

    @ShellMethod(value = "batch convert ppt/pptx to png", key = "ppt2png-batch", group = "PPTx to PNG Converter")
    public void batchConvertFiles(@ShellOption(value = "folder", defaultValue = "") String folder,
                                  @ShellOption(value = "cols", defaultValue = "3") int cols,
                                  @ShellOption(value = "space", defaultValue = "0") int space) throws Exception {
        if (StringUtils.isEmpty(folder)){
            print("Please confirm the target folder");
        } else {
            Path filePath = Paths.get(folder);
            File targetFile = new File(filePath.toUri());
            if (!targetFile.exists()){
                print("The PPTx folder dose not exists");
            }
        }

        int colsCount = cols;
        if (cols < 0 || cols > 12) {
            print("The cols should be in 1...12");
        }

        int spaceNum = space;
        if (space < 0 || space > 60) {
            print("The space between cols should be in 0...60");
        }

        AtomicInteger count = new AtomicInteger(0);
        File targetFolder = new File(folder);
        print("\nProcessing");
        folderRecursive(targetFolder, colsCount, spaceNum, count);
        print("\nJob Done");
    }

    protected void folderRecursive(File folder, int colsCount, int spaceNum, AtomicInteger count) throws Exception{
        File[] files = folder.listFiles();
        int filesCount = files.length;
        for (int i = 0; i < filesCount; i++) {
            File file = files[i];
            if (file.isFile()) {
                if (checkPPTxExtension(file.getName())){
                    process(file.getAbsolutePath(), colsCount, spaceNum);
                    print("\n" + count.addAndGet(1) + "  " + file.getName() +" completed");
                }
            } else if (file.isDirectory()){
                folderRecursive(file, colsCount, spaceNum, count);
            }
        }
    }

    /**
     * 处理一份PPTx文件
     * @param file
     * @param colsCount
     * @param spaceNum
     * @throws Exception
     */
    protected void process(String file, int colsCount, int spaceNum) throws Exception{
        Path filePath = Paths.get(file);
        File pptx = new File(filePath.toUri());
        SlideShow<?,?> shows = SlideShowFactory.create(pptx, null, true);
        List<? extends Slide<?,?>> slides = shows.getSlides();
        int slidesSize = slides.size();

        Dimension pgsize = shows.getPageSize();
        int width = (int) (pgsize.width);
        int height = (int) (pgsize.height);

        int rowsCount = (slidesSize - 1) / colsCount;
        slidesSize = rowsCount * colsCount + 1;

        /**
         * 调整slide的width和height
         */
        int slideWidth = (width - spaceNum * (colsCount - 1)) / colsCount;
        double slideRatio = slideWidth * 1.0 / width;
        int slideHeight = (int) (height * slideRatio);

        int imgFullWidth = width;
        int imgFullHeight = height + slideHeight * rowsCount + rowsCount * spaceNum;

        BufferedImage fullImg = new BufferedImage(imgFullWidth, imgFullHeight, BufferedImage.TYPE_INT_ARGB);
        Graphics2D fullGraphics = fullImg.createGraphics();
        DrawFactory.getInstance(fullGraphics).fixFonts(fullGraphics);

        fullGraphics.setColor(new Color(244, 244, 244));
        fullGraphics.fillRect(0, 0, imgFullWidth, imgFullHeight);

        // default rendering options
        fullGraphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        fullGraphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        fullGraphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        fullGraphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

        BufferedImage slideImage = null;

        for (int i = 0; i < slidesSize; i++){
            Slide<?,?> slide = slides.get(i);
            if (i == 0){
                slideImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            } else {
                slideImage = new BufferedImage(slideWidth, slideHeight, BufferedImage.TYPE_INT_ARGB);
            }

            Graphics2D slideGraphic = slideImage.createGraphics();
            DrawFactory.getInstance(slideGraphic).fixFonts(slideGraphic);

            slideGraphic.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            slideGraphic.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            slideGraphic.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            slideGraphic.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

            if (i == 0){
                slideGraphic.scale(1, 1);
            } else {
                slideGraphic.scale(slideRatio, slideRatio);
            }

            slide.draw(slideGraphic);

            if (i == 0){
                fullGraphics.drawImage(slideImage, null, 0, 0);
            } else {
                int offsetY = (i - 1) / colsCount * spaceNum + spaceNum;
                int offsetX = (i - 1) % colsCount * spaceNum;

                int ty = ((i - 1) / colsCount) * slideHeight  + height + offsetY;
                int tx = (i - 1) % colsCount * slideWidth + offsetX;

                fullGraphics.drawImage(slideImage, null, tx, ty);
            }

            slideGraphic.dispose();
            slideImage.flush();
        }

        String format = "png";
        String outname = pptx.getName().replaceFirst(".pptx?", "");
        outname = String.format(Locale.ROOT, "%1$s-%2$s.%3$s", outname, "预览", format);
        File outfile = new File(pptx.getParent(), outname);
        ImageIO.write(fullImg, "png", outfile);

        fullGraphics.dispose();
        fullImg.flush();
    }

    protected boolean checkPPTxExtension(String filename){
        int dot = filename.lastIndexOf(".");
        if (dot < 0 || dot == filename.length() - 1){
            return false;
        }

        String extension = filename.substring(dot + 1);
        if (extension.equalsIgnoreCase("ppt") || extension.equalsIgnoreCase("pptx")){
            return true;
        }

        return false;
    }

    protected void print(String message){
        System.out.println(message);
    }


//    @Component
//    public class TabValueProvider extends ValueProviderSupport {
//        private final String[] VALUES = new String[] {
//                "hello world",
//                "I'm quoting \"The Daily Mail\"",
//                "10 \\ 3 = 3"
//        };
//
//        @Override
//        public List<CompletionProposal> complete(MethodParameter parameter, CompletionContext completionContext, String[] hints) {
//            return Arrays.stream(VALUES).map(CompletionProposal::new).collect(Collectors.toList());
//        }
//    }
}
