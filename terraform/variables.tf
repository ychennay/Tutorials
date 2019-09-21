variable "region" {
    default = "us-east-1"
}

variable "default-terraform-sg" {
    default = "sg-0551d0a450fd6e498"
}

variable "incoming-sg-ports" {
    description = "Security port to accept incoming connections on."
    default = 8080
}