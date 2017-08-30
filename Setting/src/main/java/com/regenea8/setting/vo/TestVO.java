package com.regenea8.setting.vo;

public class TestVO {

	private String id, name, countrycode, district, population;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getCountrycode() {
		return countrycode;
	}

	public void setCountrycode(String countrycode) {
		this.countrycode = countrycode;
	}

	public String getDistrict() {
		return district;
	}

	public void setDistrict(String district) {
		this.district = district;
	}

	public String getPopulation() {
		return population;
	}

	public void setPopulation(String population) {
		this.population = population;
	}

	@Override
	public String toString() {
		return "TestVO [id=" + id + ", name=" + name + ", countrycode=" + countrycode + ", district=" + district
				+ ", population=" + population + "]";
	}
}
