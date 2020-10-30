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
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.example.poitldemo.domain.auto.Assets;
import com.example.poitldemo.domain.auto.AutoData;
import com.example.poitldemo.domain.auto.DetailData;
import com.example.poitldemo.policy.DetailTablePolicy;
import org.springframework.util.ClassUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.HashMap;

/**
 * <p> Title: </p>
 *
 * <p> Description: </p>
 *
 * @author: Guo Weifeng
 * @version: 1.0
 * @create: 2020/10/30 16:32
 */
@RestController
public class AutoController {
    @GetMapping("/getAuto")
    public String getAuto() throws IOException {

        String resource = ClassUtils.getDefaultClassLoader().getResource("").getPath() + "static/auto.docx";
        AutoData datas = new AutoData();
        datas.setDateTime(getNowTime());
        datas.setResponsible("");

        Assets asset = new Assets();
        asset.setArea("一区");
        asset.setNetwork("专网");
        asset.setIndex(1);
        asset.setAssetType("防火墙");
        asset.setAssetTotal(12);
        asset.setNormal(10);
        asset.setAlarm(1);
        asset.setOffline(1);


        DetailData detailData = new DetailData();
        RowRenderData detail = RowRenderData.build("1", "卫士通防火墙", "防火墙", "127.0.0.1", "专网", "在线", "50%",
            "50%", "50%", "50%", "50%", "50%", "50%", "系统运行正常");
        detailData.setDetails(Arrays.asList(detail, detail, detail));

        datas.setAssetData(Arrays.asList(asset, asset, asset));
        datas.setDetail(detailData);

        HackLoopTableRenderPolicy hackLoopTableRenderPolicy = new HackLoopTableRenderPolicy();
        Configure config = Configure.newBuilder().bind("detail", new DetailTablePolicy())
            .bind("assets", hackLoopTableRenderPolicy).build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(new HashMap<String, Object>(){
            {put("assetAuto", datas);}
        });

        template.writeToFile("out_auto.docx");

        return "Success!";
    }

    public String getNowTime() {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        return formatter.format(LocalDateTime.now());
    }
}