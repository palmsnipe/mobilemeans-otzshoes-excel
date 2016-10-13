package com.mobilemeans.otzshoes.models;

import java.util.ArrayList;
import java.util.List;

public class Shoe {
	private String name;
	private String code;
	private List<Stock> stock;
	private double priceExVAT = 0.0;
	private double retailPrice = 0.0;
	
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
	public double getPriceExVAT() {
		return priceExVAT;
	}
	public void setPriceExVAT(double priceExVAT) {
		this.priceExVAT = priceExVAT;
	}
	public double getRetailPrice() {
		return retailPrice;
	}
	public void setRetailPrice(double retailPrice) {
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
