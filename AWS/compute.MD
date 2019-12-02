# Compute

## Reserved Instances

Running a reserved instance for less than full utilization is more or less wasting money (since you are getting billed for the full 24 hours per day).

## Scheduled Reserved Instances

Scheduled Reserved Instances (Scheduled Instances) enable you to purchase capacity reservations that recur on a daily, weekly, or monthly basis, with a specified start time and duration, for a one-year term. You reserve the capacity in advance, so that you know it is available when you need it. You pay for the time that the instances are scheduled, even if you do not use them.

# EMR

## Node Types

Three types of nodes:

### Master
Manages the cluster and typically runs master components of distributed applications. Ideal for reserved/on-demand instances.

The master instance group in an EMR cluster always consists of a single node or three master nodes, so it can't scale after you initially configure it. You work with the core instance groups and task instance groups to scale out and scale in a cluster. It's possible to have a cluster with only a master node, and no core or task nodes. You must have at least one core node at cluster creation in order to scale the cluster.

### Core
Managed by the master node. Core nodes run the Data Node daemon to coordinate data storage as part of the Hadoop Distributed File System (HDFS). They also run the Task Tracker daemon and perform other parallel computation tasks on data that installed applications require (ie. YARN NodeManager daemons, Hadoop MapReduce tasks, and Spark executors). Ideal for reserved/on-demand instances.

### Task
Optional. You can use them to add power to perform parallel computation tasks on data, such as Hadoop MapReduce tasks and Spark executors. Ideal for spot instances. EMR has default functionality for scheduling YARN jobs so that running jobs don’t fail when task nodes running on Spot Instances are terminated by allowing application master processes to run only on core nodes.

## Use Cases

Customer runs an event management SaaS application that uses **Amazon EC2, Auto Scaling, Elastic Load Balancing, and Amazon RDS**. Software is installed on instances at first boot, using Puppet and Chef, which is also used to deploy software updates multiple times per week. A major software overhaul—a new, much larger version of the software—has been deployed to running EC2 instances and some of the instances are being terminated during the update process.

What actions could be taken to prevent instances from being terminated during updates?

- **Use CodeDeploy** to create an application and a deployment targeting the Auto Scaling group. Use CodeDeploy to deploy and update the application.
- **Suspend the Auto Scaling process**. Once suspended, deregister the instance from the ELB, update the application, and register it with the ELB on successful update.

Considerations:
- **Elastic Beanstalk zero downtime feature** cannot update existing applications
- The Auto-scaling process is terminating the instance since it is not responding to health checks and therefore considered unhealthy (and marked for termination).
- Using **termination protection via the console** will **only stop users, not auto-scaling groups** from deleting instances.
- Not a load balancer issue here: detaching a load balancer would stop all traffic to the instances in the ASG.
- **[EC2 autoscaling lifecycle hooks](https://docs.aws.amazon.com/autoscaling/ec2/userguide/lifecycle-hooks.html)**
- If you are using CloudFormation, you can use **[cfn-hup](https://docs.aws.amazon.com/AWSCloudFormation/latest/UserGuide/cfn-hup.html)** (The `cfn-hup` helper is a daemon that detects changes in resource metadata and runs user-specified actions when a change is detected).