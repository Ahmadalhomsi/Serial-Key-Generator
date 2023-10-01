/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package serialcodegenerator3;
import java.util.*;

/**
 *
 * @author pooqw
 */
public class Key {
    
    private String code;
    private Date StartingDate; // Starting date
    private Date ExpDate; // Expiration date
    private int devices;
    private String status;

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public int getDevices() {
        return devices;
    }

    public void setDevices(int devices) {
        this.devices = devices;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public Date getStartingDate() {
        return StartingDate;
    }

    public void setStartingDate(Date StartingDate) {
        this.StartingDate = StartingDate;
    }

    public Date getExpDate() {
        return ExpDate;
    }

    public void setExpDate(Date ExpDate) {
        this.ExpDate = ExpDate;
    }
    
    
    
}
