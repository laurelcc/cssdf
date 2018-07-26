package xin.aircloud.app.poi.ppt;

import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.xslf.util.PPTX2PNG;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.awt.image.BufferedImageOp;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Locale;

public class MainApp {

    public static void main(String[] args){

        try {
            Path filePath = Paths.get("/Users/hsoong/Desktop/PPTGROUP/[005]毕业答辩PPT模板[-吴焱鑫].ppt");
            File file = new File(filePath.toUri());
            if (file.exists() && file.isFile()){
                SlideShow<?,?> ss = SlideShowFactory.create(file, null, true);
                try {
                    List<? extends Slide<?,?>> slides = ss.getSlides();
                    int slideSize = slides.size();

                    Dimension pgsize = ss.getPageSize();
                    int width = (int) (pgsize.width);
                    int height = (int) (pgsize.height);

                    int column3Count = (int) Math.floor(slideSize / 3.0);

                    int fullWidth = width;
                    int fullHeight = height + column3Count * height / 3;

                    BufferedImage fullImg = new BufferedImage(fullWidth, fullHeight, BufferedImage.TYPE_INT_ARGB);
                    Graphics2D fullGraphics = fullImg.createGraphics();
                    DrawFactory.getInstance(fullGraphics).fixFonts(fullGraphics);

                    // default rendering options
                    fullGraphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                    fullGraphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                    fullGraphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                    fullGraphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

                    BufferedImage slideImage = null;

                    for (int i = 0; i < slideSize; i++){
                        Slide<?,?> slide = slides.get(i);

                        if (i == 0){
                            //不缩放比例绘制
                            slideImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
                        } else {
                            slideImage = new BufferedImage(width / 3, height / 3, BufferedImage.TYPE_INT_ARGB);
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
                            slideGraphic.scale(0.333333, 0.333333);
                        }

                        slide.draw(slideGraphic);

                        if (i == 0){
                            fullGraphics.drawImage(slideImage, null, 0, 0);
                        } else {
                            int ty = ((i - 1) / 3) * (int) (height * 0.333333)  + height;
                            int tx = (i - 1) % 3 * (int) (width * 0.333333);
                            fullGraphics.drawImage(slideImage, null, tx, ty);
                        }

                        slideGraphic.dispose();
                        slideImage.flush();
                    }

                    String outname = file.getName().replaceFirst(".pptx?", "");
                    outname = String.format(Locale.ROOT, "%1$s-%2$03d.%3$s", outname, 7, "png");
                    File outfile = new File(file.getParent(), outname);
                    ImageIO.write(fullImg, "png", outfile);

                    fullGraphics.dispose();
                    fullImg.flush();
                } catch (Exception e){
                    System.out.println(e.getMessage());
                } finally {
                    ss.close();
                }
            }
        } catch (Exception e){
            System.out.println(e.getMessage());
        }
    }

}
