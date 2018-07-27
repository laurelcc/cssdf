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
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
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
    public void convertSingleFile(@ShellOption(value = "--file", defaultValue = "") String file,
                                  @ShellOption(value = "--cols", defaultValue = "3") int cols,
                                  @ShellOption(value = "--space", defaultValue = "0") int space,
                                  @ShellOption(value = "--water", defaultValue = "") String water) throws Exception{
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

        if (!StringUtils.isEmpty(water)){
            Path filePath = Paths.get(water);
            File targetFile = new File(filePath.toUri());
            if (!targetFile.exists()){
                print("The wartermark file dose not exists");
            }
        }

        print("\nprocessing...");

        process(file, colsCount, spaceNum, water);

        print("\nJob Done");
    }

    @ShellMethod(value = "batch convert ppt/pptx to png", key = "ppt2png-batch", group = "PPTx to PNG Converter")
    public void batchConvertFiles(@ShellOption(value = "--folder", defaultValue = "") String folder,
                                  @ShellOption(value = "--cols", defaultValue = "3") int cols,
                                  @ShellOption(value = "--space", defaultValue = "0") int space,
                                  @ShellOption(value = "--water", defaultValue = "") String water) throws Exception {
        if (StringUtils.isEmpty(folder)){
            print("Please confirm the target folder");
            return;
        } else {
            Path filePath = Paths.get(folder);
            File targetFile = new File(filePath.toUri());
            if (!targetFile.exists()){
                print("The PPTx folder dose not exists");
                return;
            }
        }

        if (!StringUtils.isEmpty(water)){
            Path warterPath = Paths.get(water);
            File warterFile = new File(warterPath.toUri());
            if (!warterFile.exists()){
                print("The watermarks file dose not exists");
                return;
            }
        }

        int colsCount = cols;
        if (cols < 0 || cols > 12) {
            print("The cols should be in 1...12");
            return;
        }

        int spaceNum = space;
        if (space < 0 || space > 60) {
            print("The space between cols should be in 0...60");
            return;
        }

        AtomicInteger count = new AtomicInteger(0);
        File targetFolder = new File(folder);
        print("\nProcessing");
        folderRecursive(targetFolder, colsCount, spaceNum, count, water);
        print("\nJob Done");
    }

    protected void folderRecursive(File folder, int colsCount, int spaceNum, AtomicInteger count, String watermarks) throws Exception{
        File[] files = folder.listFiles();
        int filesCount = files.length;
        for (int i = 0; i < filesCount; i++) {
            File file = files[i];
            if (file.isFile()) {
                if (checkPPTxExtension(file.getName())){
                    try {
                        process(file.getAbsolutePath(), colsCount, spaceNum, watermarks);
                        print("\n" + count.addAndGet(1) + "  " + file.getName() +" completed");
                    } catch (Exception e){
                        print("\n" + file.getName() + " " + e.getMessage());
                    }
                }
            } else if (file.isDirectory()){
                folderRecursive(file, colsCount, spaceNum, count, watermarks);
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
    protected void process(String file, int colsCount, int spaceNum, String watermarks) throws Exception{
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
        fullGraphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_OFF);
        fullGraphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_SPEED);
        fullGraphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        fullGraphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_OFF);

        BufferedImage slideImage = null;

        // 画水印
        BufferedImage waterImg = null;
        if (!StringUtils.isEmpty(watermarks)){
            File waterFile = new File(watermarks);
            boolean isImgExtension = waterFile.getName().endsWith("png") ||
                    waterFile.getName().endsWith("jpg") ||
                    waterFile.getName().endsWith("jpeg");
            if (waterFile.exists() && isImgExtension){
                waterImg = ImageIO.read(waterFile);
            }
        }

        for (int i = 0; i < slidesSize; i++){
            Slide<?,?> slide = slides.get(i);
            if (i == 0){
                slideImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            } else {
                slideImage = new BufferedImage(slideWidth, slideHeight, BufferedImage.TYPE_INT_ARGB);
            }

            Graphics2D slideGraphic = slideImage.createGraphics();
            DrawFactory.getInstance(slideGraphic).fixFonts(slideGraphic);

            slideGraphic.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_OFF);
            slideGraphic.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_SPEED);
            slideGraphic.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            slideGraphic.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_OFF);

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

        // 在整张图上画水印
        if (waterImg != null){
            AffineTransform transform = new AffineTransform();
            int wih = waterImg.getHeight();
            int wiw = waterImg.getWidth();
            int waterRows = (int) (Math.ceil(fullImg.getHeight() / wih)) + 2;
            int waterCols = (int) (Math.ceil(fullImg.getWidth() / wiw)) + 2;
            for (int i = 0; i < waterRows; i++){
                for (int j = 0; j < waterCols; j++){
                    int ty = i * wih;
                    int tx = j * wiw;
                    transform.setToTranslation(tx, ty);

                    fullGraphics.drawImage(waterImg, transform, null);
                }
            }
        }

        String format = "png";
        String outname = pptx.getName().replaceFirst(".pptx?", "");
        outname = String.format(Locale.ROOT, "%1$s-%2$s.%3$s", outname, "预览", format);
        File outfile = new File(pptx.getParent(), outname);
        ImageIO.write(fullImg, format, outfile);

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

}
