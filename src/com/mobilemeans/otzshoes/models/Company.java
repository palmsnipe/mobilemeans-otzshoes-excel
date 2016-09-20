package com.mobilemeans.otzshoes.models;

import java.util.List;

public class Company {
	private String date;
	private String companyName;
	private String contactPerson;
	private String email;
	private String vatnumber;
	private String phone;
	private String storeName;
	private String delivery;
	private List<Shoe> shoes;
	
	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}
	public String getCompanyName() {
		return companyName;
	}
	public void setCompanyName(String companyName) {
		this.companyName = companyName;
	}
	public String getContactPerson() {
		return contactPerson;
	}
	public void setContactPerson(String contactPerson) {
		this.contactPerson = contactPerson;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getVatnumber() {
		return vatnumber;
	}
	public void setVatnumber(String vatnumber) {
		this.vatnumber = vatnumber;
	}
	public String getPhone() {
		return phone;
	}
	public void setPhone(String phone) {
		this.phone = phone;
	}
	public String getStoreName() {
		return storeName;
	}
	public void setStoreName(String storeName) {
		this.storeName = storeName;
	}
	public String getDelivery() {
		return delivery;
	}
	public void setDelivery(String delivery) {
		this.delivery = delivery;
	}
	public List<Shoe> getShoes() {
		return shoes;
	}
	public void setShoes(List<Shoe> shoes) {
		this.shoes = shoes;
	}
	
}
