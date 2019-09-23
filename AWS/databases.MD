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

* 20GB $\rightarrow$ 60 IOPs
* * 100GB $\rightarrow$ 300 IOPs

Maximum throughput for gp2 storage type is 1,280 Mbps (160 MBps). Suppose you have MariaDB with **page size of 16KB (.128 Mbs)**:

$$
1280Mbps\times\frac{1 page}{0.128 Mbs}\times\frac{1IO}{1 page} = 10000 IOPs
$$

You'll also need to account for how much storage to allocate

$$
10000 IOPS \times \frac{1GB}{3 IOPS} = 3333.3 GB
$$

## Credits
You can also use **burst durations** if you might need high IOPs during certain occasions but not much storage. For volumes under 1TB, you can temporarily burst to 3,000 IOPS.
$$
B = \frac{C}{3000 - 3S}
$$
$B$ is the burst duration in seconds, $C$ is the credit balance, and $S$ is the storage size in GB. A database instance is credited with an initial balance of 5,400,000 IOPs. For instance, if you have a 300GB volume and a completely full original credit balance,

$$
B = \frac{5,400,000}{3000 - 3\times300} = 2571
$$
Your database could burst for about 2,571 seconds, or roughly 40 minutes. Credit balances are replenished at a rate of 1 IOPS per second per baseline IOPS. If you have a 200GB volume database, your baseline IOPS is 600 IOPS, so to refill your credits up to the maximum of 5,400,000 credits,

$$
5400000 \times \frac{1 second}{600 IOPS}\times \frac{1 hour}{3600 seconds} = 2.5 hours
$$

## Provisioned IOPs

No concept of bursting for **io1 storage**.

The ratio of storage in gigabytes to provisioned IOPS must be at least **50:1**. For instance, 20,000 IOPs will mean at least 400GB storage.

You can allocate between 1,000 to 40,000 provisioned IOPs and between 100GB to 16TB of data. 

## DynamoDB 

NoSQL databases such as DynamoDB are ideal when you will query data based on only one attribute (column), as you can specify this as the primary key and optimize w/ an index. 


## Strongly Consistent versus Eventually Consistent

A **single strongly consistent** read of 4KB per second will consume 1 read capacity unit. An **eventual consistent** read of 2KB per second will consume 1 RCU.

### Autoscaling

1. Specify a maximum and minimum **Read** and **Write** Capacity Unit (RCU and WCUs), along with **utilization**.

* RCU: min 10, max 50
* utilization goal: 50%
* actually consume 20 -> autoscaling will set RCU at 40.

Items are distributed across partitions based on the **partition key** (AKA **hask key**)

### Composite Primary Keys

Use the combination of a **partition key** and a **sort key** to unique identify each item. Sort keys are ideal for non-unique columns that are commonly used to order records, like the `order_date` field.

## Redshift

Redshift is an **OLAP** database with **columnar storage** ideal for analytical or complex relational queries.

Types of nodes:
* **Dense compute nodes**: can store up to 326TB of data on magnetic storage.
* **Dense storage nodes**: can store up to 2PB of data on SSD (use this if you want to optimize read speeds)

Any cluster with more than one compute node includes a **leader node**.

Distribution styles:
* **EVEN**: leader node spreads data evenly across all compute notes (default style)
* **KEY**: data is distributed according to the value in a single column
* **ALL**: all tables are distirbuted to all compute nodes.

## Aurora

A **regional service** database engine compatible with MySQL (5x throughout) and PostgreSQL (3x throoughput), where compute and storage are partitioned, with a primary instance that allows reads and writes. Sync with read replicas is **synchronous**.

Two types of DB instances:
* **primary**: supports reads/writes, and is a singleton instance (only 1 primary instance)
* **replica**: also connects to the same underlying storage volume, supports only read operations (a max of 15 replica instances per Aurora DB cluster)

THe underlying storage (**cluster volume** can scale up to 64TB). Each availability zone that the cluster volume resides in will have 2 copies of the DB cluster data.

No concept of a **standby instance**.

Differences with Amazon RDS:
In RDS,
* 1 **primary instance** in **AZ1** that handles all writes (and some reads potentially)
* **Read replicas** in other Availability Zones that asynchronously replicated from the master instance

### Migration from Amazon RDS
1. Create an RDS snapshot, and then restore snapshot in Aurora
2. Set up one-way replication