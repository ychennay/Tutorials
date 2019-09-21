
# Declare available zones
data "aws_availability_zones" "available" {}

provider "aws" {
    region = "${var.region}"
}

resource "aws_security_group" "elb" {
    name = "terraform-example-elb"

    egress {
        from_port = 0
        to_port = 0
        protocol = "-1"
        cidr_blocks = ["0.0.0.0/0"]
    }

    ingress {
        from_port = 80
        to_port = 80
        protocol = "tcp"
        cidr_blocks = ["0.0.0.0/0"]
    }


}
resource "aws_launch_configuration" "example" {
  image_id = "ami-2d39803a"
  instance_type = "t2.micro"
  user_data = <<-EOF
              #! /bin/bash
              echo "Hello, world!" > index.html
              nohup busybox httpd -f -p "${var.incoming-sg-ports}" &
              EOF

  security_groups = ["${var.default-terraform-sg}"]

  lifecycle {
      create_before_destroy = true
  }
 
}

resource "aws_autoscaling_group" "terraform-asg" {
    launch_configuration = "${aws_launch_configuration.example.id}"
    availability_zones = ["${data.aws_availability_zones.available.names}"]
    min_size = 2
    max_size = 5

      
    load_balancers = ["${aws_elb.example.name}"]

    tag {
        key = "Name"
        value = "terraform-asg-example"
        propagate_at_launch = true
    }
}

resource "aws_elb" "example" {
    name = "terraform-asg-example"
    availability_zones = ["${data.aws_availability_zones.available.names}"]
    security_groups = ["${aws_security_group.elb.id}"]

    health_check {
        healthy_threshold = 2
        unhealthy_threshold = 2
        timeout = 3
        interval = 30
        target = "HTTP:${var.incoming-sg-ports}/"
        }

    listener {
        lb_port = 80
        lb_protocol = "http"
        instance_port = "${var.incoming-sg-ports}"
        instance_protocol = "http"
    }
}

output "elb_dns_name" {
    value = "${aws_elb.example.dns_name}"
}