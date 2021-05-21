package spider;

import cn.hutool.core.collection.CollectionUtil;
import org.apache.commons.io.FileUtils;
import org.apache.http.Consts;
import org.apache.http.HttpEntity;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URI;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * <dependency>
 * <groupId>org.jsoup</groupId>
 * <artifactId>jsoup</artifactId>
 * </dependency>
 */
public class DiorSpider {

    static CloseableHttpClient httpclient;
    static HttpGet httpGet;

    static int imageNum = 1;

    static String prefix = "F";

    static String zero = "0000000";

    static ExecutorService executorService = Executors.newFixedThreadPool(10);


    static {
        httpclient = HttpClients.createDefault();
        httpGet = new HttpGet();
        RequestConfig requestConfig =
                RequestConfig.custom().setConnectTimeout(20000).setConnectionRequestTimeout(20000).setSocketTimeout(20000).build();
        // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
        httpGet.setHeader("User-Agent", "Mozilla/5.0");
        httpGet.addHeader("accept", "*/*");
        httpGet.setConfig(requestConfig);
    }

    public static void main(String[] args) throws Exception {
        List<Product> products = new ArrayList<>();
        get列表页Products(products);
        get女士小型皮具Products(products);
        get男士钱包Products(products);
        get男士所有皮具列表Products(products);
        executorService.shutdown();
        while (true) {
            if (executorService.isTerminated()) {
                System.out.println("异步线程下载图片结束了！");
                break;
            }
            Thread.sleep(200);
        }
        if (CollectionUtil.isNotEmpty(products)) {
            String[] title = {"物品名称", "型号", "规格", "官网价格", "图片1", "图片2", "图片3", "图片4", "图片5", "图片6", "颜色", "材质", "描述",
                    "商品链接"};
            getHSSFWorkbook("Dior商品信息", title, products);
        }

    }


    public static void get列表页Products(List<Product> products) {
        try {
            //通过文件加载解析
            Document doc = Jsoup.parse(new File("c://Users/tt/Desktop/爬虫/dior/列表页.html"), "UTF-8");
//            Elements ulList = doc.select("div[class='page-content news-landing-page css-h14glv']");
//            Elements liList = ulList.select("div[id='prc-25-1']").select("div[class='grid-view']").select("ul[class" +
//                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
//                    " one-column legend-bottom']");

            Elements liList = doc.select("ul[class" +
                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
                    " one-column legend-bottom']");

            for (Element item : liList) {
                Product product = new Product();
                String detailUrl = "http://www.dior.cn" + item.select("div[class='product " +
                        "product-legend-bottom" +
                        " product--cdcbase']").select(
                        "a").attr("href");
                String price = item.select("div[class='product product-legend-bottom " +
                        "product--cdcbase']").select(
                        "a").select("div[class='product-legend']").select("span[class='price-line']").text().replace(
                        "¥", "").replace(",", "");
                product.setPrice(price);
                product.setDetailUrl(detailUrl);
                downPradaProductDetail(detailUrl, product);
                products.add(product);
            }
        } catch (Exception e) {
            System.out.println("获取获取商品信息失败：" + e.getMessage());
        }
    }

    public static void get女士小型皮具Products(List<Product> products) {
        try {
            //通过文件加载解析
            Document doc = Jsoup.parse(new File("c://Users/tt/Desktop/爬虫/dior/女士小型皮具.html"), "UTF-8");
//            Elements ulList = doc.select("div[class='page-content news-landing-page css-h14glv']");
//            Elements liList = ulList.select("div[id='prc-25-1']").select("div[class='grid-view']").select("ul[class" +
//                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
//                    " one-column legend-bottom']");

            Elements liList = doc.select("ul[class" +
                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
                    " one-column legend-bottom']");

            for (Element item : liList) {
                Product product = new Product();
                String detailUrl = "http://www.dior.cn" + item.select("div[class='product " +
                        "product-legend-bottom" +
                        " product--cdcbase']").select(
                        "a").attr("href");
                String price = item.select("div[class='product product-legend-bottom " +
                        "product--cdcbase']").select(
                        "a").select("div[class='product-legend']").select("span[class='price-line']").text().replace(
                        "¥", "").replace(",", "");
                product.setPrice(price);
                product.setDetailUrl(detailUrl);
                downPradaProductDetail(detailUrl, product);
                products.add(product);
            }
        } catch (Exception e) {
            System.out.println("获取获取商品信息失败：" + e.getMessage());
        }
    }

    public static void get男士所有皮具列表Products(List<Product> products) {
        try {
            //通过文件加载解析
            Document doc = Jsoup.parse(new File("c://Users/tt/Desktop/爬虫/dior/男士所有皮具.html"), "UTF-8");
//            Elements ulList = doc.select("div[class='page-content news-landing-page css-h14glv']");
//            Elements liList = ulList.select("div[id='prc-25-1']").select("div[class='grid-view']").select("ul[class" +
//                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
//                    " one-column legend-bottom']");

            Elements liList = doc.select("ul[class" +
                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
                    " one-column legend-bottom']");

            for (Element item : liList) {
                Product product = new Product();
                String detailUrl = "http://www.dior.cn" + item.select("div[class='product " +
                        "product-legend-bottom" +
                        " product--cdcbase']").select(
                        "a").attr("href");
                String price = item.select("div[class='product product-legend-bottom " +
                        "product--cdcbase']").select(
                        "a").select("div[class='product-legend']").select("span[class='price-line']").text().replace(
                        "¥", "").replace(",", "");
                product.setPrice(price);
                product.setDetailUrl(detailUrl);
                downPradaProductDetail(detailUrl, product);
                products.add(product);
            }
        } catch (Exception e) {
            System.out.println("获取获取商品信息失败：" + e.getMessage());
        }
    }

    public static void get男士钱包Products(List<Product> products) {
        try {
            //通过文件加载解析
            Document doc = Jsoup.parse(new File("c://Users/tt/Desktop/爬虫/dior/男士钱包.html"), "UTF-8");
//            Elements ulList = doc.select("div[class='page-content news-landing-page css-h14glv']");
//            Elements liList = ulList.select("div[id='prc-25-1']").select("div[class='grid-view']").select("ul[class" +
//                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
//                    " one-column legend-bottom']");

            Elements liList = doc.select("ul[class" +
                    "='grid-view-content grid-view-content--normal']").select("li[class='grid-view-element is-product" +
                    " one-column legend-bottom']");

            for (Element item : liList) {
                Product product = new Product();
                String detailUrl = "http://www.dior.cn" + item.select("div[class='product " +
                        "product-legend-bottom" +
                        " product--cdcbase']").select(
                        "a").attr("href");
                String price = item.select("div[class='product product-legend-bottom " +
                        "product--cdcbase']").select(
                        "a").select("div[class='product-legend']").select("span[class='price-line']").text().replace(
                        "¥", "").replace(",", "");
                product.setPrice(price);
                product.setDetailUrl(detailUrl);
                downPradaProductDetail(detailUrl, product);
                products.add(product);
            }
        } catch (Exception e) {
            System.out.println("获取获取商品信息失败：" + e.getMessage());
        }
    }


    public static void downPradaProductDetail(String url, Product product) throws Exception {
        httpGet.setURI(new URI(url));
        CloseableHttpResponse response = httpclient.execute(httpGet);
        // 获取响应状态码
        int statusCode = response.getStatusLine().getStatusCode();
        try {
            HttpEntity entity = response.getEntity();
            if (statusCode == 200) {
                String html = EntityUtils.toString(entity, Consts.UTF_8);
                Document doc = null;
                doc = Jsoup.parse(html);
                Elements ulList = doc.select("main[class='main']");


                Elements productInfo = ulList.select("div[class='page-content product-page-couture']").select(
                        "div[id" +
                                "='prc-23-1']").select("div[class" +
                        "='top-content-desktop-couture']").select("div[class='top-content-desktop-right']").select("div").select("div[class='top-content-desktop-sticky']");
                String productName =
                        productInfo.select("div[class='product-titles']").select("h1").select("span[class='multiline-text product-titles-title multiline-text--is-china']").text();
                product.setName(productName);
                String productNo = productInfo.select("div[class='product-titles']").select("p").text().replace("编号:", "");
                product.setProductNo(productNo);


                Elements elements = productInfo.select("div[class='product-description']").select("div[class" +
                        "='product" +
                        "-description-item']");
                if (CollectionUtil.isNotEmpty(elements) && elements.size() > 1) {
//描述
                    String desc = productInfo.select("div[class='product-description']").select("div[class='product" +
                            "-description-item']").get(0).select("div[class='product-description-item__content']").text();
                    product.setDesc(desc);


                    //规格
                    String specs = productInfo.select("div[class='product-description']").select("div[class='product" +
                            "-description-item']").get(1).select("div[class='product-description-item__content " +
                            "product-description-item__content--hidden']").text().replace("尺寸：", "");
                    product.setSpecs(specs);


                    //颜色
                    String materialColor = productInfo.select("div[class='product-titles']").select("h1").select("span" +
                            "[class" +
                            "='multiline-text product-titles-subtitle product-titles-subtitle--couture " +
                            "multiline-text--is-china']").text();
                    if (null != materialColor && materialColor.contains("色")) {
                        try {
                            String color = materialColor.substring(materialColor.indexOf("色") - 1,
                                    materialColor.indexOf("色") + 1);
                            product.setColor(color);
                            String material = materialColor.substring(materialColor.indexOf("色") + 1);
                            product.setMaterial(material);
                        } catch (Exception e) {
                            System.out.println("转换颜色、材质异常");
                        }
                    } else {
                        product.setMaterial(materialColor);
                    }
                }
                if (CollectionUtil.isNotEmpty(elements) && elements.size() == 1) {

                    String desc = productInfo.select("div[class='product-description']").select("div[class='product" +
                            "-description-item']").get(0).select("div[class='product-description-item__content']").text();
                    product.setDesc(desc);

                }


                Elements picele = ulList.select("div[class='page-content product-page-couture']").select("div[id" +
                        "='prc-23-1']").select("div[class" +
                        "='top-content-desktop-couture']").select("div[class='top-content-desktop-left']").select("ul" +
                        "[class='product-medias-grid']").select("li[class='product-medias-grid-image']");
                //获取图片
                List<String> images = new ArrayList<>();
                if (CollectionUtil.isNotEmpty(picele)) {
                    int size = Math.min(picele.size(), 6);
                    for (int i = 0; i < size; i++) {
                        String imageUrl =
                                picele.get(i).select("button[class='product-media']").select("div[class='image " +
                                        "product-media__image']").select("noscript").select(
                                        "img").attr(
                                        "src");
                        String imageName = getImageName();
//                        downloadImage(imageUrl, imageName);
                        executorService.submit(() -> {
                            downloadImage(imageUrl, imageName);
                        });
                        images.add(imageName + ".PNG");
                    }
                    product.setImages(images);
                }
                // 消耗掉实体
                EntityUtils.consume(response.getEntity());
            } else {
                // 消耗掉实体
                EntityUtils.consume(response.getEntity());
            }
        } finally {
            response.close();
        }
    }

    private static String downloadImage(String fileUrl, String fileName) {
        try {
            HttpGet httpGet = new HttpGet();
            RequestConfig requestConfig =
                    RequestConfig.custom().setConnectTimeout(20000).setConnectionRequestTimeout(20000).setSocketTimeout(20000).build();
            // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
            httpGet.setHeader("User-Agent", "Mozilla/5.0");
            httpGet.addHeader("accept", "*/*");
            httpGet.setConfig(requestConfig);
            httpGet.setURI(new URI(fileUrl));
            CloseableHttpResponse resp = httpclient.execute(httpGet);// 调用服务器接口
            String tempFileName = "c://Users/tt/Desktop/VALUE_007/VALUE_UIA/" + fileName + ".PNG";
            File temp = new File(tempFileName);
            FileUtils.copyInputStreamToFile(resp.getEntity().getContent(), temp);
            return fileName + ".PNG";
        } catch (Exception e) {
            return null;
        }
    }


    public static String downloadCompressedImage(String fileUrl, String fileName) {
        URL url = null;
        try {
            String tempFileName = "c://Users/tt/Desktop/VALUE_007/VALUE_UIA/" + fileName + ".PNG";
            url = new URL(fileUrl);
            //1.获取url的输入流 dataInputStream
            DataInputStream dataInputStream = new DataInputStream(url.openStream());
            //2.加一层BufferedInputStream
            BufferedInputStream bufferedInputStream = new BufferedInputStream(dataInputStream);
            //3.构造原始图片流 preImage
            BufferedImage preImage = ImageIO.read(bufferedInputStream);
            //4.获得原始图片的长宽 width/height
            int width = preImage.getWidth();
            int height = preImage.getHeight();
            //5.构造压缩后的图片流 image 长宽各为原来的几分之几
            BufferedImage image = new BufferedImage(width / 6, height / 6, BufferedImage.TYPE_INT_RGB);
            //6.给image创建Graphic ,在Graphic上绘制压缩后的图片
            Graphics graphic = image.createGraphics();
            graphic.drawImage(preImage, 0, 0, width / 6, height / 6, null);
            //7.为file生成对应的文件输出流
            //将image传给输出流
            FileOutputStream fileOutputStream = new FileOutputStream(tempFileName);
            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(fileOutputStream);
            //8.将image写入到file中
            ImageIO.write(image, "PNG", bufferedOutputStream);
            //9.关闭输入输出流
            bufferedInputStream.close();
            bufferedOutputStream.close();
            return fileName + ".PNG";
        } catch (IOException e) {
            return null;
        }
    }

    public static void getHSSFWorkbook(String sheetName, String[] title, List<Product> values) throws Exception {
        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        //声明列对象
        HSSFCell cell = null;
        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }
        CreationHelper createHelper = wb.getCreationHelper();
        //创建内容
        for (int i = 0; i < values.size(); i++) {
            row = sheet.createRow(i + 1);
            //将内容按顺序赋给对应的列对象
            row.createCell(0).setCellValue(values.get(i).getName());
            row.createCell(1).setCellValue(values.get(i).getProductNo());
            row.createCell(2).setCellValue(values.get(i).getSpecs());
            row.createCell(3).setCellValue(values.get(i).getPrice());
            //图片
            List<String> images = values.get(i).getImages();
            if (CollectionUtil.isNotEmpty(images)) {
                int size = Math.min(images.size(), 6);
                for (int i1 = 0; i1 < size; i1++) {
                    HSSFCell cell1 = row.createCell(i1 + 4);
                    cell1.setCellValue("VALUE_UIA\\" + images.get(i1));
                    Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.FILE);
                    hyperlink.setAddress("VALUE_UIA\\" + images.get(i1));
                    cell1.setHyperlink(hyperlink);
                }
            }
            row.createCell(10).setCellValue(values.get(i).getColor());
            row.createCell(11).setCellValue(values.get(i).getMaterial());
            row.createCell(12).setCellValue(values.get(i).getDesc());

            HSSFCell cell1 = row.createCell(13);
            cell1.setCellValue("DIOR 迪奥官网");
            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress(values.get(i).getDetailUrl());
            cell1.setHyperlink(hyperlink);

        }
        wb.write(new FileOutputStream("c://Users/tt/Desktop/VALUE_007/Dior商品信息.xls"));
    }


    static class Product {

        private String name;
        private String price;
        private String color;
        private String specs;
        private String productNo;
        private String material;
        private String desc;
        private String detailUrl;
        private List<String> images;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getPrice() {
            return price;
        }

        public void setPrice(String price) {
            this.price = price;
        }

        public String getColor() {
            return color;
        }

        public void setColor(String color) {
            this.color = color;
        }

        public String getSpecs() {
            return specs;
        }

        public void setSpecs(String specs) {
            this.specs = specs;
        }

        public String getProductNo() {
            return productNo;
        }

        public void setProductNo(String productNo) {
            this.productNo = productNo;
        }

        public String getMaterial() {
            return material;
        }

        public void setMaterial(String material) {
            this.material = material;
        }

        public String getDesc() {
            return desc;
        }

        public void setDesc(String desc) {
            this.desc = desc;
        }

        public List<String> getImages() {
            return images;
        }

        public void setImages(List<String> images) {
            this.images = images;
        }

        public String getDetailUrl() {
            return detailUrl;
        }

        public void setDetailUrl(String detailUrl) {
            this.detailUrl = detailUrl;
        }
    }

    private static String getImageName() {
        String numStr = (imageNum + "");
        StringBuilder stringBuilder = new StringBuilder(numStr);
        if (numStr.length() < 7) {
            String addStr = zero.substring(0, zero.length() - numStr.length());
            stringBuilder.insert(0, addStr).insert(0, prefix);
        } else {
            stringBuilder.insert(0, prefix);
        }
        imageNum++;
        return stringBuilder.toString();
    }


}
