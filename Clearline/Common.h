#ifndef COMMON
#define COMMON

double FilterThreePoints(double left, double centre, double right)
{
	double leftThreshold;
	double rightThreshold;

	if(centre==0) return 0;

	leftThreshold = fabs((centre - left) / centre);
	rightThreshold = fabs((centre - right) / centre);
	if((leftThreshold < 0.05) && rightThreshold < 0.05) return (left + centre + right)/3;
	return 0;
}


#endif

/*
 *  QuickSort, algorithm by C.A.R. Hoare (1960)
 *
 *  01.06.1998, implemented by Michael Neumann (neumann@s-direktnet.de)
 */

//# ifndef __QUICKSORT_HEADER__
//# define __QUICKSORT_HEADER__
//
//# include <algorithm>
//
//template <class itemType, class indexType=int>
//void QuickSort(itemType a[], indexType l, indexType r)
//{
//  static itemType m;
//  static indexType j;
//  indexType i;
//
//  if(r > l) {
    //m = a[r]; i = l-1; j = r;
    //for(;;) {
//      while(a[++i] < m);
      //while(a[--j] > m);
      //if(i >= j) break;
      //std::swap(a[i], a[j]);
    //}
    //std::swap(a[i],a[r]);
    //QuickSort(a,l,i-1);
    //QuickSort(a,i+1,r);
  //}
//}

//# endif

