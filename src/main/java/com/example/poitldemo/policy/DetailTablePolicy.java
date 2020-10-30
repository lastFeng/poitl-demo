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
package com.example.poitldemo.policy;

import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.example.poitldemo.domain.auto.DetailData;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;

/**
 * <p> Title: </p>
 *
 * <p> Description: </p>
 *
 * @author: Guo Weifeng
 * @version: 1.0
 * @create: 2020/10/30 16:55
 */
public class DetailTablePolicy extends DynamicTableRenderPolicy {

    // 数据填充位置（从0开始数）
    int detailStartRow = 3;

    @Override
    public void render(XWPFTable table, Object data) {
        if (null == data) return;
        DetailData detailData = (DetailData) data;

        List<RowRenderData> details = detailData.getDetails();
        if (null != details) {
            table.removeRow(detailStartRow);
            // 循环插入行
            for (int i = 0; i < details.size(); i++) {
                XWPFTableRow insertNewTableRow = table.insertNewTableRow(detailStartRow);
                for (int j = 0; j < 14; j++) {
                    insertNewTableRow.createCell();
                }

                // 合并单元格
                //TableTools.mergeCellsHorizonal(table, laborsStartRow, 0, 3);
                MiniTableRenderPolicy.Helper.renderRow(table, detailStartRow, details.get(i));
            }
        }
    }
}