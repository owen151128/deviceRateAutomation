package com.hpcnt.deviceRateAutomation.model;

/**
 * Created by owen151128 on 2018. 9. 6.
 */
public class Device {
    private String name;
    private String version;
    private Integer session;

    public Device(String name, String version, Integer session) {
        this.name = name;
        this.version = version;
        this.session = session;
    }

    public String getName() {
        return name;
    }

    public String getVersion() {
        return version;
    }

    public Integer getSession() {
        return session;
    }
}
