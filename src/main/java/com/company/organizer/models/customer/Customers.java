package com.company.organizer.models.customer;


        import java.util.List;

public class Customers extends DescriptionModel{

    private List<Streams> streamsList;

    public List<Streams> getStreamsList() {
        return streamsList;
    }

    public void setStreamsList(List<Streams> streamsList) {
        this.streamsList = streamsList;
    }

    @Override
    public String toString() {
        return "Customers{" +
                "streamsList=" + streamsList +
                ", rowLabels='" + rowLabels + '\'' +
                ", reportedHours='" + reportedHours + '\'' +
                ", revenue='" + revenue + '\'' +
                ", effectiveRate='" + effectiveRate + '\'' +
                '}';
    }
}
