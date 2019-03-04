package zy.poi.config;

import java.util.List;

/**
 * @author zxz
 * @date 2019-3-2
 * 合并单元格所用参数设定的类
 * mergeRowParams {
 *     最外层List:和excel中sheet数量对应，如只需处理第一个则可只设定一个，不需要处理的sheet给null
 * }
 */
public class MergeConfig {
    /**
     * 最终合并单元格所需参数集合
     */
    List<List<int []>> mergeRowParams;
    /**
     * 是否要横向合并相同的单元格
     */
    boolean isMergeColumn = false;



    public MergeConfig (){};
    public MergeConfig (List<List<int []>> mergeRowParams,boolean isMergeColumn){
        this.mergeRowParams=mergeRowParams;
        this.isMergeColumn=isMergeColumn;
    };

    public List<List<int[]>> getMergeRowParams() {
        return mergeRowParams;
    }


    public boolean isMergeColumn() {
        return isMergeColumn;
    }
}
