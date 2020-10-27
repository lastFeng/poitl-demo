/*
 * Copyright 2001-2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.example.poitldemo.controller;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.CellRenderData;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.deepoove.poi.util.BytePictureUtils;
import com.example.poitldemo.domain.Goods;
import com.example.poitldemo.domain.Labor;
import com.example.poitldemo.domain.PaymentData;
import com.example.poitldemo.util.DocUtil;
import org.springframework.util.ClassUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

/**
 * <p> Title: </p>
 *
 * <p> Description: </p>
 *
 * @author: Guo Weifeng
 * @version: 1.0
 * @create: 2020/10/27 10:56
 */
@RestController
public class WordController {

    @GetMapping("/get")
    public void getWord(HttpServletRequest request, HttpServletResponse response) {
        RowRenderData header = new RowRenderData(
            Arrays.asList(new CellRenderData(new TextRenderData("Word处理解决方案")),
                new CellRenderData(new TextRenderData("是否跨平台")),
                new CellRenderData(new TextRenderData("易用性"))));
        RowRenderData row0 = RowRenderData.build("Poi-tl", "纯Java组件，跨平台", "简单：模板引擎功能，并对POI进行了一些封装");
        RowRenderData row1 = RowRenderData.build("Apache Poi", "纯Java组件，跨平台", "简单，缺少一些功能的封装");
        RowRenderData row2 = RowRenderData.build("Freemarker", "XML操作，跨平台", "复杂，需要理解XML结构");
        RowRenderData row3 = RowRenderData.build("OpenOffice", "需要安装OpenOffice软件",
            "复杂，需要了解OpenOffice的API");
        RowRenderData row4 = RowRenderData.build("Jacob、winlib", "Windows平台", "复杂，不推荐使用");
        List<RowRenderData> tableDatas = Arrays.asList(row0, row1, row2, row3, row4);

        Map<String, Object> datas = new HashMap<String, Object>(){
            {
                put("header", "Hello");
                put("name", "POI-TL");
                put("word", "模板引擎");
                put("time", new Date());
                put("what", "Java Word模板引擎： Minimal Microsoft word(docx) templating with {{template}} in Java. It works by expanding tags in a template using values provided in a JavaMap or JavaObject.");
                put("compare", new MiniTableRenderData(header, tableDatas,
                    MiniTableRenderData.WIDTH_A4_FULL));

                put("feature", new NumbericRenderData(new ArrayList<TextRenderData>() {
                    {
                        add(new TextRenderData("Plug-in grammar, add new grammar by yourself"));
                        add(new TextRenderData(
                            "Supports word text, local pictures, web pictures, table, list, header, footer..."));
                        add(new TextRenderData(
                            "Templates, not just templates, but also style templates"));
                    }
                }));
            }
        };

        DocUtil.download(request,response,"static/template.docx","test"+ System.currentTimeMillis() + ".docx", datas);
    }

    @GetMapping("/getPolicy")
    public void getPolicy(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String resource = ClassUtils.getDefaultClassLoader().getResource("").getPath() + "static/payment_hack.docx";
        PaymentData data1 = new PaymentData();
        PaymentData data2 = new PaymentData();

        data1.setSubtotal("8000");
        data1.setTax("600");
        data1.setTransform("120");
        data1.setOther("250");
        data1.setUnpay("6600");
        data1.setTotal("总共：7200");

        List<Goods> goods = new ArrayList<>();
        Goods good = new Goods();
        good.setCount(4);
        good.setName("墙纸");
        good.setDesc("书房卧室");
        good.setDiscount(1500);
        good.setPrice(400);
        good.setTax(new Random().nextInt(10) + 20);
        good.setTotalPrice(1600);
        good.setPicture(new PictureRenderData(24, 24, ".png", BytePictureUtils.getUrlBufferedImage("http://deepoove.com/images/icecream.png")));
        goods.add(good);
        goods.add(good);
        goods.add(good);
        goods.add(good);
        data1.setGoods(goods);

        List<Labor> labors = new ArrayList<>();
        Labor labor = new Labor();
        labor.setCategory("油漆工");
        labor.setPeople(2);
        labor.setPrice(400);
        labor.setTotalPrice(1600);
        labors.add(labor);
        labors.add(labor);
        labors.add(labor);
        data1.setLabors(labors);

        data2.setSubtotal("8002");
        data2.setTax("602");
        data2.setTransform("122");
        data2.setOther("252");
        data2.setUnpay("6602");
        data2.setTotal("总共：7202");

        data2.setGoods(goods);

        HackLoopTableRenderPolicy hackLoopTableRenderPolicy = new HackLoopTableRenderPolicy();
        Configure config = Configure.newBuilder().bind("goods", hackLoopTableRenderPolicy)
            .bind("labors", hackLoopTableRenderPolicy).build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(new HashMap<String, Object>() {
            {
                put("datas", Arrays.asList(data1, data2));
            }
        });
        template.writeToFile("out_example_payment_hack"+ System.currentTimeMillis() +".docx");
    }
}