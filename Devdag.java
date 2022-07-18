package dagautomation;

import com.mongodb.BasicDBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import java.io.*;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;

public class DagDevelopment {

    WebDriver driver;

    String rawVendorName_AL = "";
    String rawSourceName_AL = "";
    String visited_date_AL = "";

    // TestBase testbase_obj = new TestBase();
    int count = 0;
    int serialNo = 1;
    String modified_producturl;

    // String rawVendorName = "vitalsource";
    int port = 61336;

    // String rawSourceName = "flexoffers";

    // String visited_date = "2020-08-21T00:00:00.000Z";


    @Test
    public void mongoDBQualityTesting() {

        ArrayList<Object[]> dagdata = new ArrayList<Object[]>();

       /* dagdata.add(new String[] { "backmarket", "awin", "2021-06-21T00:00:00.000Z" });
        dagdata.add(new String[] { "myprotein", "awin", "2021-06-21T00:00:00.000Z" });
        dagdata.add(new String[] { "viator", "awin", "2021-06-21T00:00:00.000Z" });*/
        
 //      dagdata.add(new String[] { "pedometersusa", "scraper", "2021-12-07T00:00:00.000Z" });
 //      dagdata.add(new String[] { "eugenetoyandhobby", "scraper", "2021-12-07T00:00:00.000Z" });
 //      dagdata.add(new String[] { "pedometersusa", "scraper", "2021-12-07T00:00:00.000Z" });
 //      dagdata.add(new String[] { "portableone", "scraper", "2021-12-07T00:00:00.000Z" });

 //      dagdata.add(new String[] { "birthdaydirect", "scraper", "2021-12-08T00:00:00.000Z" });
 //      dagdata.add(new String[] { "radartoys", "scraper", "2021-12-08T00:00:00.000Z" });
        
        dagdata.add(new String[] { "swap", "linkshare", "2022-06-20T00:00:00.000Z" });
//        dagdata.add(new String[] { "rosegal", "scraper", "2022-06-10T00:00:00.000Z" });
//       dagdata.add(new String[] { "spongelle", "scraper", "2022-06-08T00:00:00.000Z" });
//        dagdata.add(new String[] { "hunterbellnyc", "scraper", "2022-06-08T00:00:00.000Z" });
//        dagdata.add(new String[] { "bobssportschalet", "scraper", "2021-11-19T00:00:00.000Z" });
/*        dagdata.add(new String[] { "heartratemonitorsusa", "scraper", "2021-11-19T00:00:00.000Z" });
        dagdata.add(new String[] { "toywiz", "scraper", "2021-11-19T00:00:00.000Z" });
        dagdata.add(new String[] { "birthdayexpress", "scraper", "2021-11-19T00:00:00.000Z" });
        dagdata.add(new String[] { "marathonsports", "scraper", "2021-11-19T00:00:00.000Z" });*/

        MongoClient mongoClientObj = new MongoClient("127.0.0.1", port);
        MongoDatabase db = mongoClientObj.getDatabase("pdcdb1");

        for (Object[] md : dagdata) {
            for (int i = 0; i < md.length; i++) {

                if (i == 0) {
                    // System.out.println("rawVendorName--="+md[i]);
                    rawVendorName_AL = (String) md[i];
                    System.out.println("#######################" + rawVendorName_AL);
                }
                if (i == 1) {
                    // System.out.println("rawSourceName--="+md[i]);
                    rawSourceName_AL = (String) md[i];
                    System.out.println("#######################" + rawSourceName_AL);
                }
                if (i == 2) {
                    // System.out.println("visited_date--="+md[i]);
                    visited_date_AL = (String) md[i];
                    System.out.println("#######################" + visited_date_AL);
                }

            }

            for (String name : db.listCollectionNames()) {
                System.out.println(name);
            }

            BasicDBObject whereQuery = new BasicDBObject();
            // whereQuery.put("raw_vendor", rawVendorName);
            whereQuery.put("raw_vendor", rawVendorName_AL);

//		Instant instant = Instant.parse("2020-05-18T00:00:00.000Z");
            // Instant instant = Instant.parse(visited_date);
            Instant instant = Instant.parse(visited_date_AL);
            Date timestamp = Date.from(instant);

            Document query = new Document();
//		query.put("raw_source", rawSourceName);
//		query.put("raw_vendor", rawVendorName);
            query.put("raw_source", rawSourceName_AL);
            query.put("raw_vendor", rawVendorName_AL);
            query.put("visited", new BasicDBObject("$gt", timestamp));

//		FindIterable<Document> mydbrecords = db.getCollection("product_listing").find(whereQuery);
            FindIterable<Document> mydbrecords = db.getCollection("product_listing").find(query);

            MongoCursor<Document> iterator = mydbrecords.iterator();

            FileInputStream f = null;
            try {
                f = new FileInputStream("C:\\Users\\vmtes\\Desktop\\Dag Dev.xlsx");
            } catch (FileNotFoundException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            XSSFWorkbook workbook = null;
            try {
                workbook = new XSSFWorkbook(f);
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            String sheetName = "MongoSheet";
            XSSFSheet sheet = workbook.getSheet(sheetName);

            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy HH-mm-ss");
            LocalDateTime now = LocalDateTime.now();
            System.out.println(dtf.format(now).toString());
            String datetime = dtf.format(now).toString();

            // Create a blank sheet
//		XSSFSheet sheet = workbook.createSheet(rawVendorName + datetime);
            XSSFSheet sheet1 = workbook.createSheet(rawVendorName_AL + datetime);

            ArrayList<Object[]> data = new ArrayList<Object[]>();
            data.add(new String[] { "S.No", "id", "productUrl", "imageUrl", "affiliateUrl", "affilUrl", "productName",
                    "brandName", "skuCondition", "raw_vendor", "raw_source", "raw_url", "productDescription", "Data",
                    "Mongo", "Site", "Remarks" });

            while (iterator.hasNext()) {
                Document doc = iterator.next();

                // String upc = doc.getString("upc");
                // if(upc != null && !upc.isEmpty())
                // {
                // System.out.println("Upc: " + upc);
                // }
                // else
                // {
                // System.out.println("No Data");
                // }
                int num = serialNo++;
                String sNo = String.valueOf(num);
                Document attributes = (Document) doc.get("attributes");
                Document raw_attributes = (Document) doc.get("raw_attributes");
                // String size = attributes.getString("size");
                // System.out.println("****** Size :" + size);

                String id = doc.getObjectId("_id").toHexString();
                String raw_source = doc.getString("raw_source");
                String raw_vendor = doc.getString("raw_vendor");
                // String raw_url = doc.getString("raw_url");
                String productUrl = doc.getString("productUrl");
                String imageUrl = doc.getString("imageUrl");

                String raw_url;
                if ((doc.getString("raw_url")) == null || (doc.getString("raw_url")).isEmpty()) {
                    raw_url = "Null";

                } else {
                    raw_url = doc.getString("raw_url");
                }

                String affiliateUrl;
                if ((doc.getString("affiliateUrl")) == null || (doc.getString("affiliateUrl")).isEmpty()) {
                    affiliateUrl = "Null";

                } else {
                    affiliateUrl = doc.getString("affiliateUrl");
                }

                String affilUrl;
                if ((doc.getString("affilUrl")) == null || (doc.getString("affilUrl")).isEmpty()) {
                    affilUrl = "Null";
                } else {
                    affilUrl = doc.getString("affilUrl");
                }
                String productName = doc.getString("productName");
                // String quality = "Null";
                String site = "Null";

                Object remarks = null;
                if ((doc.getString("raw_source")) == null || (doc.getString("raw_source")).isEmpty()) {
                    String raw_source_remarks1 = "\n ** raw_source is missing ";
                    remarks = remarks + raw_source_remarks1;

                }

                if ((doc.getString("raw_vendor")) == null || (doc.getString("raw_vendor")).isEmpty()) {
                    String raw_vendor_remarks1 = "\n **  raw_vendor is missing ";
                    remarks = remarks + raw_vendor_remarks1;

                }

                if ((doc.getString("productUrl")) == null || (doc.getString("productUrl")).isEmpty()) {
                    String productUrl_remarks1 = "\n ** productUrl is missing ";
                    remarks = remarks + productUrl_remarks1;

                }

                if ((doc.getString("productName")) == null || (doc.getString("productName")).isEmpty()) {
                    String productName_remarks1 = "\n ** productName is missing ";
                    remarks = remarks + productName_remarks1;

                }

                String productDescription;
                if ((doc.getString("productDescription")) == null || (doc.getString("productDescription")).isEmpty()) {
                    productDescription = "";

                } else {
                    productDescription = doc.getString("productDescription");
                }

                Object extradata = null;
                if ((doc.getString("sku")) == null || (doc.getString("sku")).isEmpty()) {
                    String sku_remarks1 = "\n** sku is missing ";
                    remarks = remarks + sku_remarks1;
                } else {
                    String sku = "\n sku : " + doc.getString("sku");
                    extradata = extradata + sku;
                }

                if ((doc.getString("upc")) == null || (doc.getString("upc")).isEmpty()) {
                    String upc_remarks1 = "\n ** upc is missing ";
                    remarks = remarks + upc_remarks1;
                } else {
                    String upc = "\n upc : " + doc.getString("upc");
                    extradata = extradata + upc;
                }

                if ((doc.getString("mpn")) == null || (doc.getString("mpn")).isEmpty()) {
                    String mpn_remarks1 = "\n ** mpn is missing ";
                    remarks = remarks + mpn_remarks1;

                } else {
                    String mpn = "\n mpn : " + doc.getString("mpn");
                    extradata = extradata + mpn;
                }

                String brandName;
                if ((doc.getString("brandName")) == null || (doc.getString("brandName")).isEmpty()) {

                    brandName = "";

                    String brandName_remarks1 = "\n ** brandName is missing ";
                    remarks = remarks + brandName_remarks1;

                } else {
                    brandName = doc.getString("brandName");
                }

                if ((doc.getString("imageUrl")) == null || (doc.getString("imageUrl")).isEmpty()) {
                    String imageUrl_remarks1 = "\n ** imageUrl is missing ";
                    remarks = remarks + imageUrl_remarks1;

                }

                if ((doc.getString("affiliateUrl")) == null
                        || (doc.getString("affiliateUrl")).isEmpty() && (doc.getString("affilUrl")) == null
                        || (doc.getString("affilUrl")).isEmpty()) {
                    String affiliateUrl_remarks1 = "\n ** affiliateUrl is missing ";
                    remarks = remarks + affiliateUrl_remarks1;

                }

                Object mongo = null;
                if ((doc.getString("price")) == null || (doc.getString("price")).isEmpty()) {
                    String price_remarks1 = "\n ** price is missing ";
                    remarks = remarks + price_remarks1;

                    String price = "\n price : ";
                    mongo = mongo + price;

                } else {
                    String price = "\n price : " + doc.getString("price");
                    mongo = mongo + price;
                }

                if ((doc.getString("salePrice")) == null || (doc.getString("salePrice")).isEmpty()) {
                    String salePrice_remarks1 = "\n ** salePrice is missing ";
                    remarks = remarks + salePrice_remarks1;

                    String salePrice = "\n salePrice : ";
                    mongo = mongo + salePrice;

                } else {
                    String salePrice = "\n salePrice : " + doc.getString("salePrice");
                    mongo = mongo + salePrice;
                }

                if ((doc.getString("basePrice")) == null || (doc.getString("basePrice")).isEmpty()) {
                    String basePrice_remarks1 = "\n ** basePrice is missing ";
                    remarks = remarks + basePrice_remarks1;

                    String basePrice = "\n basePrice : ";
                    mongo = mongo + basePrice;

                } else {
                    String basePrice = "\n basePrice : " + doc.getString("basePrice");
                    mongo = mongo + basePrice;

                }

                String skuCondition = null;
                if ((doc.getString("skuCondition")) == null || (doc.getString("skuCondition")).isEmpty()) {
                    String skuCondition_remarks1 = "\n ** skuCondition is missing ";
                    remarks = remarks + skuCondition_remarks1;
                } else {
                    skuCondition = doc.getString("skuCondition");
                }

//		    if ((doc.getString("catPrimary")) == null || (doc.getString("catPrimary")).isEmpty() && (doc.getString("raw_category")) == null || (doc.getString("raw_category")).isEmpty())
//		   if ((doc.getString("catPrimary")) == null || (doc.getString("catPrimary")).isEmpty())
                if ((doc.getString("catPrimary")).isEmpty() && (doc.getString("raw_category")).isEmpty()) {
                    String catPrimary_remarks1 = "\n ** catPrimary and raw_category is missing ";
                    remarks = remarks + catPrimary_remarks1;
                } else {
                    String catPrimary = "\n catPrimary : " + doc.getString("catPrimary");
                    extradata = extradata + catPrimary;
                    String catSecondary = "\n catSecondary : " + doc.getString("catSecondary");
                    extradata = extradata + catSecondary;
                    String raw_category = "\n raw_category : " + doc.getString("raw_category");
                    extradata = extradata + raw_category;
                }

                if (raw_attributes.isEmpty()) {
                    String empty = "\n**Raw_attributes**\n Empty ";
                    mongo = mongo + empty;
                } else {
                    if ((raw_attributes.getString("color")) == null || (raw_attributes.getString("color")).isEmpty()) {
                        String color = "\n**Raw_attributes**\n color : ";
                        mongo = mongo + color;
                    } else {
                        String color = "\n**Raw_attributes**\n color : " + raw_attributes.getString("color");
                        mongo = mongo + color;
                    }

                    if ((raw_attributes.getString("age")) == null || (raw_attributes.getString("age")).isEmpty()) {
                        String age = "\n age : ";
                        mongo = mongo + age;
                    } else {
                        String age = "\n age : " + raw_attributes.getString("age");
                        mongo = mongo + age;

                    }

                    if ((raw_attributes.getString("gender")) == null
                            || (raw_attributes.getString("gender")).isEmpty()) {
                        String gender = "\n gender : ";
                        mongo = mongo + gender;
                    } else {
                        String gender = "\n gender : " + raw_attributes.getString("gender");
                        mongo = mongo + gender;

                    }

                    if ((raw_attributes.getString("size")) == null || (raw_attributes.getString("size")).isEmpty()) {
                        String size = "\n size : ";
                        mongo = mongo + size;
                    } else {
                        String size = "\n size : " + raw_attributes.getString("size");
                        mongo = mongo + size;

                    }
                }

                if ((attributes.getString("color")) == null || (attributes.getString("color")).isEmpty()) {
                    String color = "\n**Attributes**\n color : ";
                    mongo = mongo + color;
                } else {
                    String color = "\n**Attributes**\n color : " + attributes.getString("color");
                    mongo = mongo + color;

                }

                if ((attributes.getString("age")) == null || (attributes.getString("age")).isEmpty()) {
                    String age = "\n age : ";
                    mongo = mongo + age;
                } else {
                    String age = "\n age : " + attributes.getString("age");
                    mongo = mongo + age;

                }

                if ((attributes.getString("gender")) == null || (attributes.getString("gender")).isEmpty()) {
                    String gender = "\n gender : ";
                    mongo = mongo + gender;
                } else {
                    String gender = "\n gender : " + attributes.getString("gender");
                    mongo = mongo + gender;

                }

                if ((attributes.getString("size")) == null || (attributes.getString("size")).isEmpty()) {
                    String size = "\n size : ";
                    mongo = mongo + size;
                } else {
                    String size = "\n size : " + attributes.getString("size");
                    mongo = mongo + size;

                }

                if ((doc.getString("productKey")) == null || (doc.getString("productKey")).isEmpty()) {
                    String productKey_remarks1 = "\n ** productKey is missing ";
                    remarks = remarks + productKey_remarks1;
                } else {
                    String productKey = "\n productKey : " + doc.getString("productKey");
                    extradata = extradata + productKey;
                }

                String remarks2 = "\n Null";

                if (!productUrl.startsWith("https://")) {
                    modified_producturl = "https://" + productUrl;
                } else {
                    String site_url = productUrl;
                }

                String site_url = modified_producturl;
                String browser_type = "Chrome";

                data.add(new String[] { sNo, id, productUrl, imageUrl, affiliateUrl, affilUrl, productName, brandName,
                        skuCondition, raw_vendor, raw_source, raw_url, productDescription, (String) extradata, (String) mongo, site,
                        (String) remarks });

                affiliateUrl = "";
                affilUrl = "";
                mongo = "* Price *";
                remarks = "Remarks";
                extradata = "Data";
                count++;
            }

            // Iterate over data and write to sheet
            int rownum = 0;
            for (Object[] mongodata : data) {
                Row row = sheet1.createRow(rownum++);

                int cellnum = 0;
                for (Object obj : mongodata) {
                    Cell cell = row.createCell(cellnum++);
                    if (obj instanceof String)
                        cell.setCellValue((String) obj);
                    else if (obj instanceof Double)
                        cell.setCellValue((Double) obj);
                    else if (obj instanceof Integer)
                        cell.setCellValue((Integer) obj);
                }
            }
            try {
                // Write the workbook in file system
                FileOutputStream out = new FileOutputStream(new File("C:\\Users\\vmtes\\Desktop\\Dag Dev.xlsx"));
                workbook.write(out);
                out.close();
                System.out.println("Excel Mongo.xlsx has been created successfully");
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                try {
                    workbook.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }

    }

    @AfterTest
    public void afterTest() {
        System.out.println("Total Count :" + count);
    }

    public void verifyTextPresent(String value) {
        driver.getPageSource().contains(value);
    }
}
