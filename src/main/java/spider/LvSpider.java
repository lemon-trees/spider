package spider;

import cn.hutool.core.collection.CollectionUtil;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
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
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * <dependency>
 * <groupId>org.jsoup</groupId>
 * <artifactId>jsoup</artifactId>
 * </dependency>
 */
public class LvSpider {

    static CloseableHttpClient httpclient;
    static HttpGet httpGet;

    static int imageNum = 1;

    static String prefix = "C";

    static String zero = "0000000";

    static String file_path = "c://Users/tt/Desktop/VALUE_004";

    static String image_folder = "VALUE_RIA";

    static String image_tyep = ".PNG";

    static String women_all_bag_url = "https://www.louisvuitton.cn/zhs-cn/women/handbags/all-handbags/_/N-1ouyuai" +
            "?page=";
    static String men_all_bag_url = "https://www.louisvuitton.cn/zhs-cn/men/bags/all-bags/_/N-nstx58?page=";

    static ExecutorService executorService = Executors.newFixedThreadPool(40);


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
        Map<String, Product> products = new HashMap<>();
        for (int i = 1; i < 40; i++) {
            downProduct(women_all_bag_url, i, products);
        }
        for (int i = 1; i < 40; i++) {
            downProduct(men_all_bag_url, i, products);
        }

        executorService.shutdown();
        while (true) {
            if (executorService.isTerminated()) {
                System.out.println("异步线程下载图片结束了！");
                break;
            }
            Thread.sleep(200);
        }
        if (CollectionUtil.isNotEmpty(products.values())) {
            String[] title = {"物品名称", "型号", "规格", "官网价格", "图片1", "图片2", "图片3", "图片4", "图片5", "图片6", "颜色", "材质", "描述",
                    "商品链接"};
            getHSSFWorkbook("prada商品信息", title, new ArrayList<>(products.values()));
        }
    }


    public static void downProduct(String url, int page, Map<String, Product> products) throws Exception {
        // 需要爬取商品信息的网站地址
        url = url + page;
        // 动态模拟请求数据
        httpGet.setURI(new URI(url));
        // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
        httpGet.setHeader("user-agent", "Mozilla/5.0");
        try {
            CloseableHttpResponse response = httpclient.execute(httpGet);
            // 获取响应状态码
            int statusCode = response.getStatusLine().getStatusCode();
            HttpEntity entity = response.getEntity();
            if (statusCode == 200) {
                String html = EntityUtils.toString(entity, Consts.UTF_8);
                Document doc = Jsoup.parse(html);
                Elements div = doc.select("div[class='lv-paginated-list lv-category__grid']");
                Elements liList = div.select("div[class='lv-paginated-list lv-category__grid']").select("ul[class='lv" +
                        "-list']").select("li");
                if (CollectionUtil.isEmpty(liList)) {
                    return;
                }
                for (int i = 0; i < liList.size(); i++) {
                    Element item = liList.get(i);
                    Elements info = item.select("div[class='lv-product-card -compact-large']").select("div" +
                            "[class='lv-product-card__wrap']")
                            .select("div[class='lv-product-card__info-wrapper']").select("div[class='lv-product" +
                                    "-card__info']").select("span");
                    Product product = new Product();
                    String productNo = info.attr("id").replace("product-", "");
                    if (StringUtils.isEmpty(productNo)) {
                        continue;
                    }
                    product.setProductNo(productNo);

                    String name = info.select("a").text();
                    product.setName(name);

                    String detailUrl = "https://www.louisvuitton.cn" + info.select("a").attr("href");
                    product.setDetailUrl(detailUrl);
                    products.put(productNo, product);
                    downProductDetail(detailUrl, product);
                }
                // 消耗掉实体
                EntityUtils.consume(response.getEntity());
            } else {
                // 消耗掉实体
                EntityUtils.consume(response.getEntity());
            }
            response.close();
        } catch (Exception e) {
            System.out.println("请求列表失败：" + e.getMessage());
        } finally {

        }
    }


    public static void downProductDetail(String url, Product product) throws Exception {
        httpGet.setURI(new URI(url));
        CloseableHttpResponse response = httpclient.execute(httpGet);
        // 获取响应状态码
        int statusCode = response.getStatusLine().getStatusCode();
        try {
            HttpEntity entity = response.getEntity();
            if (statusCode == 200) {
                String html = EntityUtils.toString(entity, Consts.UTF_8);
                Document doc = Jsoup.parse(html);
                Elements ulList = doc.select("div[class='lv-product__wrap']");


                Elements detail = ulList.select("div[class='lv-product__details']");

                String price = detail.select("div[class='lv-product__price-stock']").text().replace("¥", "").replace(",", "");
                product.setPrice(price);


                Elements colorAndMaterial = detail.select("div[class='lv-product-variations']")
                        .select("button[class='lv-product-variation-selector list-label-l " +
                                "lv-product-variations__selector']");
                if (CollectionUtil.isNotEmpty(colorAndMaterial)) {
                    String typeName = colorAndMaterial.get(0)
                            .select("span[class='lv-product-variation-selector__title -text-is-medium']").text();

                    if ("颜色".equals(typeName)) {
                        String color = colorAndMaterial.get(0)
                                .select("span[class='lv-product-variation-selector__value']").text();
                        product.setColor(color);
                    }

                    if (colorAndMaterial.size() > 1) {
                        typeName = colorAndMaterial.get(1)
                                .select("span[class='lv-product-variation-selector__title -text-is-medium']").text();
                        if ("材质".equals(typeName)) {
                            String material = colorAndMaterial.get(1)
                                    .select("span[class='lv-product-variation-selector__value']").text();
                            product.setMaterial(material);
                        }
                    }
                }


                //规格
                String specs = detail.select("div[class='lv-product-features']")
                        .select("div[class='lv-expandable-panel -minimal-height']")
                        .select("div[class='lv-expandable-panel__content']").select("bdo").text();
                if (null != specs) {
                    specs = specs.replace("(长度 x 高 x 宽)", "").replace("厘米", "cm");
                    product.setSpecs(specs);
                }

                String desc = detail.select("div[class='lv-product-description']").text();
                String detailSesc = detail.select("div[class='lv-product-features']")
                        .select("div[class='lv-expandable-panel -minimal-height']")
                        .select("div[class='lv-expandable-panel__content']").text();
                desc = desc + detailSesc;
                product.setDesc(desc);

                //获取图片
                Elements picele = ulList.select("div[class='lv-product__primary']").select("div[class='lv-product__primary-wrap']")
                        .select("div[class='lv-product-visual']")
                        .select("ul[class='lv-product-visual-module lv-list lv-product-visual__module-desktop']").select("li");
                List<String> images = new ArrayList<>();

                int size = Math.min(picele.size(), 6);

                for (int i = 0; i < size; i++) {
                    String imageUrl = picele.get(i).select("button[class='lv-product-visual-module__button']")
                            .select("div[class='lv-smart-picture lv-responsive-picture lv-product-visual-module__image -fit-contain']")
                            .select("noscript").select("img").attr("src");
                    String imageName = getImageName();
                    executorService.submit(() -> {
                        downloadImage(imageUrl, imageName);
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
            String tempFileName = file_path + "/" + image_folder + "/" + fileName + ".PNG";
            File temp = new File(tempFileName);
            FileUtils.copyInputStreamToFile(resp.getEntity().getContent(), temp);
            return fileName + ".PNG";
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
        URL url;
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

            int widthNew = width;
            int heightNew = height;

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
            cell1.setCellValue("LouisVuitton 路易威登官网");
            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
            hyperlink.setAddress(values.get(i).getDetailUrl());
            cell1.setHyperlink(hyperlink);

        }
        wb.write(new FileOutputStream(file_path + "/商品信息.xls"));
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
