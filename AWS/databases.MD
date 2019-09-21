# Databases

## Important Metrics

* **RPO (recovery point objective)**: maximum period of acceptable data loss (*how much data are you okay losing in the event of a failure?*)

* **RTO (recovery time objective)**: maximum acceptable time to recover data and resume normal operations after failure.

## Point in Time Recovery

Enabling **point-in-time recovery** generates database change logs written to S3 each 5 minutes. 

## Service Levels

AWS gaurantees service levels for provisioned IOPS within a **10%** range for 99.9% of the year (ie., any time except for 2 hours and 45 minutes a year).

## General-Purpose SSD

For each GB of data allocated for storage, RDS will provide 3 IOPs, up to 10,000 IOPS:

## Provisioned IOPs

No concept of bursting for **io1 storage**.

The ratio of storage in gigabytes to provisioned IOPS must be at least **50:1**. For instance, 20,000 IOPs will mean at least 400GB storage.

You can allocate between 1,000 to 40,000 provisioned IOPs and between 100GB to 16TB of data. 

## DynamoDB Autoscaling

1. Specify a maximum and minimum **Read** and **Write** Capacity Unit (RCU and WCUs), along with **utilization**.

* RCU: min 10, max 50
* utilization goal: 50%
* actually consume 20 -> autoscaling will set RCU at 40.

Items are distributed across partitions based on the **partition key** (AKA **hask key**)

### Composite Primary Keys

Use the combination of a **partition key** and a **sort key** to unique identify each item.