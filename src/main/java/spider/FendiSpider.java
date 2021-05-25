package spider;

import cn.hutool.core.collection.CollectionUtil;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
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
public class FendiSpider {

    static CloseableHttpClient httpclient;
    static HttpGet httpGet;

    static int imageNum = 1;

    static String prefix = "J";

    static String zero = "0000000";

    static String file_path = "c://Users/tt/Desktop/VALUE_011";

    static String image_folder = "VALUE_YIA";

    static String image_tyep = ".PNG";

    static ExecutorService executorService = Executors.newFixedThreadPool(30);


    static {
        httpclient = HttpClients.createDefault();
        httpGet = new HttpGet();
        // 模拟浏览器浏览（user-agent的值可以通过浏览器浏览，查看发出请求的头文件获取）
        httpGet.setHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36");
    }

    public static void main(String[] args) throws Exception {

        List<Product> products = new ArrayList<>();

        downProduct("c://Users/tt/Desktop/爬虫/fendi/女士皮夹.json", products);
        downProduct("c://Users/tt/Desktop/爬虫/fendi/女士手袋.json", products);
        downProduct("c://Users/tt/Desktop/爬虫/fendi/男士皮夹.json", products);
        downProduct("c://Users/tt/Desktop/爬虫/fendi/男士手袋.json", products);

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

    public static String readJsonFile(String fileName) {
        String jsonStr = "";
        try {
            File jsonFile = new File(fileName);
            FileReader fileReader = new FileReader(jsonFile);
            Reader reader = new InputStreamReader(new FileInputStream(jsonFile), "utf-8");
            int ch = 0;
            StringBuffer sb = new StringBuffer();
            while ((ch = reader.read()) != -1) {
                sb.append((char) ch);
            }
            fileReader.close();
            reader.close();
            jsonStr = sb.toString();
            return jsonStr;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void downProduct(String url, List<Product> products) throws Exception {
        String test = readJsonFile(url);
        JSONObject jobj = JSON.parseObject(test);
        JSONArray jsonArray = jobj.getJSONObject("data").getJSONArray("items");
        for (Object o : jsonArray) {
            Product product = new Product();
            JSONObject ob = (JSONObject) o;
            product.setName(ob.getString("name") + ob.getString("shortDesc"));
            product.setDesc(ob.getString("description"));
            product.setPrice(ob.getString("price").replace("￥", "").replace(",", ""));
            product.setProductNo(ob.getString("sku"));
            String detailUrl = "https://www.fendi.cn/rest/default/V1/applet/product/" +
                    ob.getString("sku") +
                    "?version=7765a92ee7e1b31628f9917d9a6aacac";
            product.setDetailUrl(ob.getString("url"));
            JSONObject chiil = (JSONObject) ob.getJSONArray("childProducts").get(0);
            product.setColor(chiil.getString("colorCn"));

            downPradaProductDetail(detailUrl, product);

            products.add(product);
        }


    }


    public static void downPradaProductDetail(String url, Product product) throws Exception {
        httpGet.setURI(new URI(url));
        CloseableHttpResponse response = httpclient.execute(httpGet);
        try {
            // 获取响应状态码
            int statusCode = response.getStatusLine().getStatusCode();
            HttpEntity entity = response.getEntity();
            if (statusCode == 200) {
                String json = EntityUtils.toString(entity, Consts.UTF_8);
                JSONObject jobj = JSON.parseObject(json).getJSONObject("data");
                JSONArray attr = jobj.getJSONArray("childProducts");

                if (null != attr) {
                    JSONObject info = ((JSONObject) attr.get(0)).getJSONObject("specificMorer");
                    String material = info.getString("composition");
                    product.setMaterial(material);


                    String spec = "";
                    String height = info.getString("height");
                    if (StringUtils.isNotBlank(height)) {
                        if (height.contains(",") && height.lastIndexOf(",") == height.length() - 1) {
                            height = height.substring(0, height.lastIndexOf(","));
                        }
                        spec += height + "x";
                    }
                    String depth = info.getString("depth");
                    if (StringUtils.isNotBlank(depth)) {
                        if (depth.contains(",") && depth.lastIndexOf(",") == depth.length() - 1) {
                            depth = depth.substring(0, depth.lastIndexOf(","));
                        }
                        spec += depth + "x";
                    }
                    String length = info.getString("length");
                    if (StringUtils.isNotBlank(length)) {
                        if (length.contains(",") && length.lastIndexOf(",") == length.length() - 1) {
                            length = length.substring(0, length.lastIndexOf(","));
                        }
                        spec += length + "x";
                    }
                    if (StringUtils.isNotBlank(spec) && spec.contains("x")) {
                        spec = spec.substring(0, spec.lastIndexOf("x")).replace("Cm","");
                        product.setSpecs(spec + "cm");
                    }

                }


                JSONArray imagesArray = jobj.getJSONArray("images");
                if (null != imagesArray) {
                    List<String> images = new ArrayList<>();
                    List<String> imagesUrl = new ArrayList<>();
                    for (Object o : imagesArray) {
                        JSONObject image = (JSONObject) o;
                        imagesUrl.add(image.getString("url"));
                        imagesUrl.add(image.getString("tiny_url"));
                        if (imagesUrl.size() > 6) {
                            break;
                        }
                    }
                    for (String item : imagesUrl) {
                        String imageName = getImageName();
                        executorService.submit(() -> {
                            downloadCompressedImage(item, imageName);
                        });
                        images.add(imageName + ".PNG");
                    }
                    product.setImages(images);
                }
                EntityUtils.consume(response.getEntity());
            } else {
                // 消耗掉实体
                EntityUtils.consume(response.getEntity());
            }
        } catch (Exception e) {
            System.out.println("获取商品详情失败" + e.getMessage());

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

            int widthNew = width / 2;
            int heightNew = height / 2;

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
            cell1.setCellValue("FENDI 芬迪官网");
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
