// C program to multiply two square matrices. 
#include <stdio.h> 
#include <stdlib.h>
#include <sys/time.h>
#include <math.h>
#define I 400
#define J 400
#define K 400
#define RANDOM_SIZE 10

int operations;
float A[I][K], B[K][J], C[I][J];
void multiply() 
{ 
    int i, j, k; 
    for (i = 0; i < I; i++) 
    { 
        for (j = 0; j < J; j++) 
        { 
            C[i][j] = 0; 
            for (k = 0; k < K; k++) 
                C[i][j] += A[i][k] * B[k][j];
                operations++; 
        } 
    } 
} 

void make_matrices()
{
	printf("Creating random array\n");
	int i,j,k;
	for (i=0; i<I; i++) {
		for (k=0; k<K; k++) {
			A[i][k] = (float)rand()/RANDOM_SIZE;
		}
	}
	for (k=0; k<K; k++) {
		for (j=0; j<J; j++) {
			B[k][j] = (float)rand()/RANDOM_SIZE;
		}
	}   
}
int main() 
{  
    int i,j;

   	// timer structs
    struct	timeval ts, te, td;
    float tser, tpar, diff;

    make_matrices(); // generate random matrices
    
    gettimeofday(&ts, NULL); // start timer
    multiply(); 
    gettimeofday(&te, NULL); // end timer
    printf("Result matrix is \n"); 
    for (i = 0; i < I; i++) 
    { 
        for (j = 0; j < J; j++) 
           printf("%f ", C[i][j]); 
        printf("\n"); 
    }
   printf("A total of %d floating point operations were performed.\n", operations); 
    timersub(&ts, &te, &td);
    tser = fabs(td.tv_sec+(float)td.tv_usec/1000000.0);
    printf("Time : %.2f sec \n\n", tser );  
    float gflops;

   gflops = (operations / 1000000000.0) / tser;
   printf("GFLOPS: %.3f\n", gflops);

    return 0; 
} 

