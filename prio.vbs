option explicit
'------------------------------------------------
'priority queue class in a self-dimensioning array
'AGV 2019
'a private out_of_order function must be adapted according to the kind of data in the queue and the order required
'--------------------------------------
Class prio_queue
  private size
  private q 
  
  'adapt this function to your data
  private function out_of_order(father,son):out_of_order= (father>son): end function  
 
 function peek
    peek=q(1)
 end function
 
 property get qty
    qty=size
 end property  

property get isempty 
    isempty=(size=0)
end property
 
 function remove
    dim x
    x=q(1)
    q(1)=q(size)
    size=size-1
    sift_down
    remove=x
 end function
 
 sub add (x)
    size=size+1
    if size>ubound(q) then redim preserve q(ubound(q)+100)
    q(size)=x
    sift_up
 end sub   
 
 Private sub swap (i,j)
    dim x
    x=q(i):q(i)=q(j):q(j)=x
  end sub   
  
  private sub sift_up
    dim h,p
    h=size:
    p=h\2
    while out_of_order(q(p),q(h)) and h>1
       swap h,p
       h=p 
       p=h\2
    wend       
  end sub  
  
  private sub sift_down
  dim p,h
  p=1
  do
    if p>=size then exit do
    h =p*2 
    if h >size then exit do
    if h+1<=size then if  out_of_order(q(h),q(h+1)) then h=h+1
    if out_of_order(q(p),q(h)) then swap h,p
    p=h      
  loop
 end sub   
   

  'Al instanciar objeto con New
  Private Sub Class_Initialize(  )
      redim q(100)
     size=0
  End Sub

  'When Object is Set to Nothing
  Private Sub Class_Terminate(  )
	    erase q
  End Sub
End Class
'-------------------------------------
'test program 
'---------------------------------

dim queue,i,o,n,ercnt
set queue=new prio_queue
wscript.echo "Adding 2000 random inputs to the queue"
for i=1 to 2000
   queue.add(cint(rnd*10000))
next

wscript.echo  "Done. Using .qty and .peek methods: " & queue.qty() &" items in queue. Item "&  queue.peek()& " is at the top." & vbcrlf

wscript.echo "Removing 2000 items from the queue and giving an error if one of them is smaller than the previous one"
o=-99999999:i=1:ercnt=0
while not queue.isempty()
   n=queue.remove()
   if o>n then 
        ercnt=ercnt+1
        wscript.echo  "error at item "& i & " Previous was " & o & " and present is " & n
   end if     
   o= n:i=i+1
wend
wscript.echo "Finished. " & ercnt & " items out of order. " & i-1  & " items removed. "& n & " was the last one." 
set queue= nothing
    

   
