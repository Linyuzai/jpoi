package com.github.linyuzai.jpoi.excel.value;

import org.apache.poi.ss.usermodel.ClientAnchor;

public class ClientAnchorValue implements SupportClientAnchorValue {

    private Padding padding;
    private Location location;
    private ClientAnchor.AnchorType anchorType;

    public ClientAnchorValue() {
    }

    public ClientAnchorValue(Padding padding, Location location, ClientAnchor.AnchorType anchorType) {
        this.padding = padding;
        this.location = location;
        this.anchorType = anchorType;
    }

    @Override
    public Padding getPadding() {
        return padding;
    }

    public void setPadding(Padding padding) {
        this.padding = padding;
    }

    @Override
    public Location getLocation() {
        return location;
    }

    public void setLocation(Location location) {
        this.location = location;
    }

    @Override
    public ClientAnchor.AnchorType getAnchorType() {
        return anchorType;
    }

    public void setAnchorType(ClientAnchor.AnchorType anchorType) {
        this.anchorType = anchorType;
    }
}
