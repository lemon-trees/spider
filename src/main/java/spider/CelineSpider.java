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
public class CelineSpider {

    static CloseableHttpClient httpclient;
    static HttpGet httpGet;

    static int imageNum = 1;

    static String prefix = "D";

    static String zero = "0000000";

    static String file_path = "c://Users/tt/Desktop/VALUE_005";

    static String image_folder = "VALUE_SIA";

    static String image_tyep = ".PNG";

    static ExecutorService executorService = Executors.newFixedThreadPool(20);


    static {
        httpclient = HttpClients.createDefault();
        httpGet = new HttpGet();
        // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
        httpGet.setHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36");
    }

    public static void main(String[] args) throws Exception {

        List<Product> products = new ArrayList<>();

        String[] womenHandBagCategorys = {"新品", "triomphe帆布", "triomphe-刺绣织物", "triomphe", "belt-bag", "classic", "luggage", "16", "sangle", "cabas手袋"};
        String url = "https://www.celine.com/zhs-cn/celine%E5%A5%B3%E5%A3%AB/%E6%89%8B%E8%A2%8B/";
        for (String womenHandBagCategory : womenHandBagCategorys) {
            downProduct(url + womenHandBagCategory + "/", products);
        }
        String[] womenWalletCategorys = {"新品", "essentials", "皮具配饰", "triomphe帆布"};
        url = "https://www.celine.com/zhs-cn/celine%E5%A5%B3%E5%A3%AB/%E5%B0%8F%E7%9A%AE%E5%85%B7";
        for (String womenWalletCategory : womenWalletCategorys) {
            downProduct(url + womenWalletCategory + "/", products);
        }


        String[] menCategorys = {"新品", "邮差包", "单肩背包", "cabas手袋", "商务与旅行手袋", "背包"};
        url = "https://www.celine.com/zhs-cn/celine%E7%94%B7%E5%A3%AB/%E6%89%8B%E8%A2%8B/";
        for (String menCategory : menCategorys) {
            downProduct(url + menCategory + "/", products);
        }

        String[] menWalletCategorys = {"新品", "钱包", "卡包", "电子产品配饰", "商务与旅行手袋", "背包"};
        url = "https://www.celine.com/zhs-cn/celine%E7%94%B7%E5%A3%AB/%E5%B0%8F%E7%9A%AE%E5%85%B7/";
        for (String menWalletCategory : menWalletCategorys) {
            downProduct(url + menWalletCategory + "/", products);
        }

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
            getHSSFWorkbook("商品信息", title, products);
        }
    }

    public static void downProduct(String url, List<Product> products) throws Exception {
        // 动态模拟请求数据
        httpGet.setURI(new URI(url));
        // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
        httpGet.setHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36");
        CloseableHttpResponse response = httpclient.execute(httpGet);
        // 获取响应状态码
        int statusCode = response.getStatusLine().getStatusCode();
        try {
            HttpEntity entity = response.getEntity();
            if (statusCode == 200) {
                String html = EntityUtils.toString(entity, Consts.UTF_8);
                Document doc = Jsoup.parse(html);
                Elements ulList = doc.select("ul[class='o-listing-grid']");
                Elements liList = ulList.select("li");
                for (Element item : liList) {
                    Product product = new Product();

                    String productNo = item.select("div[class='m-product-listing']").attr("data-id");
                    product.setProductNo(productNo);

                    Elements productEle = item.select("div[class='m-product-listing']").select("a");
                    String productName = productEle.select("div[class='m-product-listing__wrapper']").select("div" +
                            "[class" +
                            "='m-product-listing__meta']").select("div[class='m-product-listing__meta-title f-body']").text();
                    product.setName(productName);
                    String price = productEle.select("div[class='m-product-listing__wrapper']").select("div" +
                            "[class" +
                            "='m-product-listing__meta']").select("p[class='m-product-listing__meta-price f-body']").text().replace("CNY", "");
                    product.setPrice(price);
                    String color = productEle.select("div[class='m-product-listing__wrapper']").select("div" +
                            "[class" +
                            "='m-product-listing__meta']").select("div[class='m-product-listing__meta-title f-body']").select("span[class='a11y']").text().replace(";", "");
                    product.setColor(color);

                    String detailUrl = "https://www.celine.com" + productEle.attr("href");
                    product.setDetailUrl(detailUrl);
                    downPradaProductDetail(detailUrl, product);
                    products.add(product);
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
                Elements ulList = doc.select("div[class='o-product__product']");
                String desc = ulList.select("div[class='o-product__meta']").select("form[class='o-form" +
                        "']").select("div[class='o-product__description o-body-copy']").select("div[class='a-text " +
                        "f-body']").select("p").html();

                if (null != desc) {
                    if (desc.contains("<br>")) {
                        String[] info = desc.split("<br>");
                        String specs = info[0];
                        if (specs.contains("(") && specs.contains(")")) {
                            specs = specs.substring(specs.indexOf("(") + 1, specs.indexOf(")"));
                        }
                        product.setSpecs(specs);
                        String material = info[1];
                        product.setMaterial(material);
                    }
                    product.setDesc(desc.replace("<br>", ","));
                }


                //获取图片
                Elements picele = ulList.select("div[class='o-product__imgs']").select("ul[class='m-thumb-carousel " +
                        "m-thumb-carousel--prime m-thumb-carousel--square" +
                        "']").select("li");
                List<String> images = new ArrayList<>();
                for (Element item : picele) {
                    String imageUrl =
                            item.select("button[class='m-thumb-carousel__img']").select("img").attr(
                                    "data-src-zoom");
                    String imageName = getImageName();
                    executorService.submit(() -> {
                        downloadCompressedImage(imageUrl, imageName);
                    });
                    images.add(imageName + ".PNG");
                }
                product.setImages(images);
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
//            URL url = new URL(fileUrl);
////            String tempFileName = "c://Users/tt/Desktop/VALUE_010/VALUE_XIA/" + fileName + ".webp";
//            String tempFileName = file_path + image_folder + fileName + ".webp";
//            File temp = new File(tempFileName);
//            FileUtils.copyURLToFile(url, temp);
//            return fileName + ".webp";

            HttpGet httpGet = new HttpGet();
            RequestConfig requestConfig =
                    RequestConfig.custom().setConnectTimeout(20000).setConnectionRequestTimeout(20000).setSocketTimeout(20000).build();
            // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
            httpGet.setHeader("User-Agent", "Mozilla/5.0");
            httpGet.addHeader("accept", "*/*");
            httpGet.setConfig(requestConfig);
            httpGet.setURI(new URI(fileUrl));
            CloseableHttpResponse resp = httpclient.execute(httpGet);// 调用服务器接口
            String tempFileName = file_path + "/" + image_folder + "/" + fileName + ".jpg";
            File temp = new File(tempFileName);
            FileUtils.copyInputStreamToFile(resp.getEntity().getContent(), temp);
            return fileName + ".jpg";
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * https://blog.csdn.net/monitor1394/article/details/6087583
     *
     * @param fileUrl
     * @param fileName
     * @return
     */
    public static String downloadCompressedImage(String fileUrl, String fileName) {
        URL url = null;
        try {
            String tempFileName = file_path + "/" + image_folder + "/" + fileName + ".PNG";
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

            int widthNew = width / 6;
            int heightNew = height / 6;

            //5.构造压缩后的图片流 image 长宽各为原来的几分之几
            BufferedImage image = new BufferedImage(widthNew, heightNew, BufferedImage.TYPE_INT_ARGB);
            //6.给image创建Graphic ,在Graphic上绘制压缩后的图片
            Graphics2D g2d = image.createGraphics();
            image = g2d.getDeviceConfiguration().createCompatibleImage(widthNew, heightNew, Transparency.TRANSLUCENT);
            g2d.dispose();
            g2d = image.createGraphics();

            Image from = preImage.getScaledInstance(widthNew, heightNew, preImage.SCALE_AREA_AVERAGING);
            g2d.drawImage(from, 0, 0, null);
            g2d.dispose();


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
                for (int i1 = 0; i1 < images.size(); i1++) {
                    HSSFCell cell1 = row.createCell(i1 + 4);
                    cell1.setCellValue(image_folder + "\\" + images.get(i1));
                    Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.FILE);
                    hyperlink.setAddress(image_folder + "\\" + images.get(i1));
                    cell1.setHyperlink(hyperlink);
                }
            }
            row.createCell(10).setCellValue(values.get(i).getColor());
            row.createCell(11).setCellValue(values.get(i).getMaterial());
            row.createCell(12).setCellValue(values.get(i).getDesc());


            HSSFCell cell1 = row.createCell(13);
            cell1.setCellValue("SELINE 思琳官网");
            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress(values.get(i).getDetailUrl());
            cell1.setHyperlink(hyperlink);

        }
        wb.write(new FileOutputStream(file_path + "/商品信息.xls"));
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
