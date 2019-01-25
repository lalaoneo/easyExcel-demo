package com.lori.easyExceldemo.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.HashMap;
import java.util.Map;

public class ExcelPersonListener extends AnalysisEventListener {

    private Integer successCount = 0;

    private Integer failCount = 0;

    private Map<Integer,String> map = new HashMap<>();

    public Map<Integer, String> getMap() {
        return map;
    }

    public Integer getSuccessCount() {
        return successCount;
    }

    public Integer getFailCount() {
        return failCount;
    }

    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        System.out.println("当前行: "+analysisContext.getCurrentRowNum());
        if ((analysisContext.getCurrentRowNum()-analysisContext.getCurrentSheet().getHeadLineMun()) < 1000){
            System.out.println(o);
            successCount++;
        }else {
            failCount++;
            map.put(analysisContext.getCurrentRowNum(),"只能接受1000行数据");
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("解析完成！！！");
    }
}
