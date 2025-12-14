package com.changgou.order.service;

import com.changgou.order.pojo.Order;
import com.changgou.order.pojo.OrderItem;
import entity.Result;

import java.util.List;

public interface CartService {

    /**
     * 加入购物车
     */
    void add(Integer num, Long id, String username);

    List<OrderItem> list(String username);
}
