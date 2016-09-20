package com.mobilemeans.otzshoes.models;

import java.util.ArrayList;
import java.util.List;

public class Shoe {
	private String name;
	private String code;
	private List<Stock> stock;
	private float priceExVAT;
	private float retailPrice;
	
	public Shoe() {
		this.stock = new ArrayList<Stock>();
	}
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getCode() {
		return code;
	}
	public void setCode(String code) {
		this.code = code;
	}
	public List<Stock> getStock() {
		return stock;
	}
	public void setStock(List<Stock> stock) {
		this.stock = stock;
	}
	public void addStock(Stock stock) {
		this.stock.add(stock);
	}
	public float getPriceExVAT() {
		return priceExVAT;
	}
	public void setPriceExVAT(float priceExVAT) {
		this.priceExVAT = priceExVAT;
	}
	public float getRetailPrice() {
		return retailPrice;
	}
	public void setRetailPrice(float retailPrice) {
		this.retailPrice = retailPrice;
	}
	
	public int getTotalStock() {
		int total = 0;
		
		for (Stock size : stock) {
			total += size.getQuantity();
		}
		
		return total;
	}
	
}
