package com.seetech.automation.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import org.testng.Assert;

public class ConfigManager {
	private static Map<String, Properties> configs = new HashMap<String, Properties>();
	public static String configFileName = System.getProperty("user.dir")+File.separator+"ConfigData"+File.separator+"config.properties";

	private static void loadConfig(String name)
	{
		Properties prop = new Properties();
		try
		{
			InputStream inputStream = new FileInputStream(configFileName);
			prop.load(inputStream);
			inputStream.close();
		}
		catch (IOException e)
		{
			Assert.fail(e.getMessage());
			return;
		}
		configs.put(name, prop);
	}

	public static Properties getConfig(String name)
	{
		if(!configs.containsKey(name))
		{
			loadConfig(name);
		}
		return configs.get(name);
	}

	public static String getProperty( String propertyName)
	{
		return getConfig(configFileName).getProperty(propertyName).toString();
	}

	public static void main(String[] args){

		System.out.println( getProperty("url"));

	}

}
