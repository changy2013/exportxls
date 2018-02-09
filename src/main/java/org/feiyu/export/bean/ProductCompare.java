package org.feiyu.export.bean;

import java.util.Comparator;

/**
 * @ Author : yangchang@ctrip.com
 * @ Desc ：
 * @ Date : Created in 2017/12/29 14:55
 * @ Modified By ：
 */
public class ProductCompare implements Comparator<GoonProduct> {
    @Override
    public int compare(GoonProduct o1, GoonProduct o2) {
        return o1.getSkuCode().compareTo(o2.getSkuCode());
    }
}
