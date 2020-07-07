	program beltcast
	include 'mgimp.inc'
	
	integer iconv,isp,iter,ites3,itin,itcool,item,ihint
	double precision resref,avt,check,heatin,heatout
	double precision totflux,genth
c        write(6,*)' entering readddata'
	call readdata(file1,lres,nx,ny,ihint)
c        write(6,*)' entering calcgeom'
c  reads all input data ( including property arrays )

	call calcgeom(nx,ny)
c  sets up geometric information

	if ( lq ) then
c          write(6,*)' entering fluxint'
	  call fluxint(nx)
	end if
c  interpolates heat flux data onto grid
c        write(6,*)' entering initv'
	call initv(file1,lres,nx,ny)
c  intialise all variables

	iconv=0
	if ( ls ) then
	  isp=0
	  write(6,*)' speed will be adjusted periodically'
	  write(6,789)speed*60.
789       format(2x,' initial speed is ',f6.2,' m/min')
	else
	  isp=1
	end if
	iter=0
	if ( ls ) niter=niter*4
	ites3=0
	itin=tinlet
	resref=data(itin,7)*data(itin,10)*speed*tinlet

	if ( lres ) then
c          write(6,*)' entering calcaux'
	  call calcaux(nx,ny)
	end if

c  main loop starts here

	do while ( (iconv.eq.0 .or. isp.eq.0) .and. iter.le. niter )
c               ! outer convergence loop
	  iter=iter+1

	  if ( iter.eq.1 )then
	    if ( lsol ) then
c              write(6,*)' entering calccoeffs'
	      call calccoeffs(nx,ny,iter)
	    else
c              write(6,*)' entering calccoeffsp'
	      call calccoeffsp(nx,ny,iter)
	    end if
	  end if


	  if ( lsol ) then
c            write(6,*)' entering solve'
	    call solve(nx,ny)
	  else
c            write(6,*)' entering solvep'
	    call solvep(nx,ny)
	  end if

c          write(6,*)' defining max and min'          
	  do ix=1,nx
	  do iy=1,ny
	    if ( tem(ix,iy) .lt. tcool ) tem(ix,iy)=tcool
	    if ( tem(ix,iy) .gt. tinlet ) tem(ix,iy)=tinlet
	  end do
	  end do

c          write(6,*)' entering updatef'
	  call updatef(nx,ny,ihint)
c  calculate new fraction solids

c          write(6,*)' entering calcaux'
	  call calcaux(nx,ny)

c  calculate auxiliary info ( e.g. h.t.c.'s and shrinkage )

	  if ( ls .and. mod(iter,nsweep) .eq. 0 ) then
c            write(6,*)' call upds ',iter,nsweep,ls,isp
	    call updatespeed(avt,nx,ny,isp,ites3)
	    write(6,457)avt,texit
	    write(6,456)speed*60.
	  end if
456     format(2x,' speed is changed to ',f6.3,' m/min ')
457     format(2x,' calc exit temp is ',f6.1,
     1 ' -- desired is ',f6.1)

c          write(6,*)' entering update'
	  call update(nx,ny)
c  update temperatures + frac sols ( + relaxation if necess )

	  if ( lsol ) then
c            write(6,*)' entering calccoeffs'
	    call calccoeffs(nx,ny,iter)
	  else
c            write(6,*)' entering calccoeffsp'
	    call calccoeffsp(nx,ny,iter)
	  end if

c          write(6,*)' entering calconv'
	  call calconv(nx,ny,iconv,iter,resref)
c  convergence criterion ( + residuals )

	  if ( iter.eq.1 ) then
	    open(92,file='mon.1',status='unknown')
	    open(93,file='mon.2',status='unknown')
	    open(94,file='mon.3',status='unknown')
	    open(95,file='chk.1',status='unknown')
	  end if
	  write(92,*)iter,tem(1,1),fr(1,1)
	  write(93,*)iter,tem(1,ny),fr(1,ny)
	  write(94,*)iter,tem(nx/4,ny/2),fr(nx/4,ny/2)
	  check=p(1,1,5)*tem(1,1)-p(1,1,2)*tem(1,2)-
     1          sou(1,1)-p(1,1,4)*tem(1,0)-p(1,1,1)*tem(0,1)
	  write(95,*)iter,check,p(1,1,5)*tem(1,1)
	end do

	if ( iter .ge. niter ) then
	  write(6,*)' maximum number of iterations reached'
	  write(6,*)' check parameters'
	end if

	itcool=tcool
	itin=tinlet
	heatin=speed*gy*data(itcool,10)
     1         *data(itin,8)
	totflux=0.0d0
	do ix=1,nx
	  totflux=totflux+qflux(ix,1)*dxu(ix)
	  totflux=totflux+qflux(ix,2)*dxu(ix)
	end do
	heatout=0.0d0
	do iy=1,ny
	  item=tem(nx,iy)
	  genth=data(item,8)+(tem(nx,iy)-item)*data(item,9)
	  heatout=heatout+genth*dyv(iy)
     1            *data(itcool,10)
	end do
	heatout=speed*heatout
	write(6,*)
	write(6,994)heatin
	write(6,995)heatout
	write(6,996)totflux
	write(6,997)heatin-heatout-totflux
	write(6,*)
	avt=0.0d0
	do iy=1,ny
	  avt=avt+tem(nx,iy)*dyv(iy)
	end do
	avt=avt/yv(ny)
	write(6,998)avt
	write(6,*)
	write(6,999)speed*60.
	write(6,*)' writing out results'
	call dumpsol(nx,ny)
	write(6,*)

	stop
994     format(2x,'net energy in is       ',f12.0,
     1 ' watts/unit width')
995     format(2x,'net energy out is      ',f12.0,
     1 ' watts/unit width')
996     format(2x,'int. flux from belt is ',f12.0,
     1 ' watts/unit width')
997     format(2x,'global error is        ',f12.0,
     1 ' watts/unit width')

998     format(2x,' average exit temperature is ',f6.2)
999     format(2x,' for casting speed of ',f5.2,' m/min')
	end
	subroutine readdata(file1,lres,nx,ny,ihint)
	include 'mgimp.inc'
	integer ihint,itf,itl,j23,iocheck,ji,j22
	double precision ascf94,bscf94,cscf94,dscf94,escf94

	open(2,file='INPUT.DAT',status='old')

	read(2,*)nx,ny
	if ( nx.gt.mx .or. ny.gt.my ) then
	  write(6,*)' parameter statement in mgimp.inc not large enough'
	  write(6,*)' mx is ',mx,' and my is ',my
	  write(6,*)' needed mx is ',nx,' and my is ',ny
	  write(6,*)' program stopping'
	  stop
	end if


	read(2,*)ialloy
c alloy number


c	stop

	read(2,*)gx,gy
c length of caster ; thickness of caster

	read(2,*)glat
c latent heat ( can be got from DATA file )


	read(2,*)lq,ls,lsol
c logicals for heat flux given (T) or htc given (F)
c and speed given (T) or exit temp given (F)
c lsol=t ==> elliptic solver ; else parabolic



	read(2,*)tinlet,speed,texit,tcool,taper
c inlet temp, inlet speed, exit temp, coolant temp, taper


	read(2,*)shrnk,fstart,fsqueeze
c volume shrinkage ( on solid. ), frac sol when gap forms, frac sol where
c squeeze principle applies  ( frac sol refers to integrated frac. sol. )

	read(2,*)hmax,fmax,hmin
c max. h.t.c. , h=hmax when  frac sol < fmax, min. h.t.c.

	read(2,*)hpar1,hpar2,hpar3
	hpar2=dsqrt(gy/hpar2)

	read(2,*)hpar4,hpar5,hpar6
c six parameters for use in h.t.c. correlation

	read(2,*)liter,dtfals,difcut
c no of internal iterations on solver, false time step, upwinding param

	read(2,*)relaxh,relaxf,nsweep
c linear relaxation on temp. , and fraction solid , number of sweeps before
c speed updated

	read(2,*)gtes4,gfrac,gclose,niter
c amount to change speed by to match exit temp; frac of inlet enth
c for convergence criteria; closeness to desired exit temperature ;
c maximum number of sweeps possible

	read(2,*)lres
	if ( lres ) then
	  read(2,'(a20)')file1
	else
	  read(2,*)  ! not really necessary
	end if
c lres=T for restart ; file1 is name of restart file

	read(2,*)

	if ( lq ) then
c		read(2,*)
	  read(2,*)nflux
c number of measured flux data

c	stop

	  do i=1,nflux
	    read(2,*)xc(i),gmeasflux(i,1),gmeasflux(i,2)
c x pos of head flux data; bottom data ; top data
	  end do
	end if
	close(2)

c read in property data

	IHINT=1000            !  no of intervals in enthalpy range
	ITF=tcool-1           !  min temp for prop range
	ITL=tinlet+1          !  max temp for temp range
	open(79,file='DATA.DAT',status='old')
	open(78,file='ENTH.DAT',status='old')
	read(79,*)
	j23=itf
1003    read(79,*,END=1004,IOSTAT=iocheck)ji,(data(j23,j),j=1,11)
	j23=j23+1
	if ( j23 .gt. nd ) then
	  write(6,*)' no of property points is greater than permitted'
	  write(6,*)' probably casting temperature is greater than'
	  write(6,*)' permitted maximum of ',nd,' deg C'
	  write(6,*)' check data --- if o.k. then change parameter'
	  write(6,*)' nd in include file to casting temperature'
	  write(6,*)' program stopping'
	  stop
	end if
	goto 1003
1004    continue
c        write(6,*)' no of rows in property data is ',j23-1-itf
	CLOSE(79)
	j23=0
	READ(78,*)
c        do ji=0,nd
c          read(78,*)j23,(enth(ji,j22),j22=1,5)
c        end do
	 iocheck=0
1005    read(78,*,END=1006,IOSTAT=iocheck)ji,ascf94,bscf94,cscf94,
     1                                    dscf94,escf94
	enth(j23,1)=ascf94
	enth(j23,2)=bscf94
	enth(j23,3)=cscf94
	enth(j23,4)=dscf94
	enth(j23,5)=escf94
c        write(6,*)ji,j23,iocheck
	j23=j23+1
	goto 1005
1006    continue
	if ( j23 .eq. ( ihint+1 ) ) then
	  write(6,*)' property data files read in O.K.'
	else
	  write(6,*)' enth.dat file does not contain ',ihint,
     1  ' enthalpy intervals'
	write(6,*)' no of intervals is ',j23-1 
	  write(6,*)' -- program will not work properly'  
	  write(6,*)' seek help -- program stopping'
	  stop
	end if
	CLOSE(78)
	return
	end
	subroutine calcgeom(nx,ny)
	include 'mgimp.inc'
	double precision gx1,dx,dy

	xu(0)=0.0d0
	gx1=gx*1.2  ! make computational domain 20% longer than caster
	dx=gx1/nx
	do i=1,nx
	  xu(i)=xu(i-1)+dx
	end do
	do i=1,nx
	  xg(i)=(xu(i-1)+xu(i))/2.0
	  dxu(i)=xu(i)-xu(i-1)
	  if ( i.gt.1 ) dxg(i)= xg(i)-xg(i-1)
	end do

	yv(0)=0.0d0
	dy=gy/ny
	do i=1,ny
	  yv(i)=yv(i-1)+dy
	end do
	do i=1,ny
	  yg(i)=(yv(i-1)+yv(i))/2.0
	  dyv(i)=yv(i)-yv(i-1)
	  if ( i.gt.1 ) dyg(i)= yg(i)-yg(i-1)
	end do
	return
	end
	subroutine initv(file1,lres,nx,ny)
	include 'mgimp.inc'
	integer ierr,nx1,ny1
	logical ltemp
C
C
c
c   default frac solids to zero
	ltemp=lres
	if ( lres ) then
	  open(48,file=file1,err=20,iostat=ierr,status='old')
	  goto 21
20        continue
	  write(6,*)' error in opening restart file'
	  write(6,*)' iostat is ',ierr
	  write(6,*)' default values will be taken'
	  ltemp=.FALSE.
	  goto 22
21        continue
	  read(48,*,err=23)nx1,ny1
	  goto 24
23        write(6,*)' nothing in file ',file1
	  write(6,*)' default numbers taken'
	  ltemp=.FALSE.
	  goto 22
24        continue
	  if ( nx1 .eq. nx .and. ny1 .eq. ny ) then
	  else
	    write(6,*)' restart file dimensions not consistent'
	    write(6,*)' nx is ',nx1,' in restart file'
	    write(6,*)' while nx is ',nx,'in present run'
	    write(6,*)' ny is ',ny1,' in restart file'
	    write(6,*)' while ny is ',ny,'in present run'
	    write(6,*)' default values will be taken'
	    ltemp=.FALSE.
	  end if
22        continue
	  close(48)
	end if
	lres=ltemp
	if ( lres ) then
	  open(48,file=file1)
	  do iy=1,ny
	    read(48,*)(tem(ix,iy),ix=1,nx)
	    read(48,*)(fr(ix,iy),ix=1,nx)
	  end do
	  close(48)
	  do ix=1,nx
	  do iy=1,ny
	    temold(ix,iy)=tem(ix,iy)
	    frold(ix,iy)=fr(ix,iy)
	  end do
	  end do
	else
	  do ix=1,nx
	    do iy=1,ny
	      tem(ix,iy)=tinlet
	      temold(ix,iy)=tinlet
	      fr(ix,iy)=0.0d0
	      frold(ix,iy)=0.0d0
	    end do
	  end do
	end if
	ix=0
	do iy=1,ny
	  fr(ix,iy)=0.0d0   ! inlet fraction solid ( essential 0
	  frold(ix,iy)=0.0d0   ! inlet fraction solid ( essential 0
	end do
c
c
c  it is essential that the temps at ix=0 are put at the inlet temperature
c  and that the temperatures at iy=0 and iy=ny+1 are put equal to tcool
c  temps at ix=nx+1 are irrelevant
c  these comments apply to restarts as well

	do ix=nx+1,nx+1
	  do iy=1,ny
	    tem(ix,iy)=tem(nx,iy)
	    temold(ix,iy)=tem(nx,iy)
	  end do
	end do

	ix=0
	do iy=1,ny
	  tem(ix,iy)=tinlet
	  temold(ix,iy)=tinlet
	end do

	iy=0
	do ix=0,nx+1
	  tem(ix,iy)=tcool
	  temold(ix,iy)=tcool
	end do

	iy=ny+1
	do ix=0,nx+1
	  tem(ix,iy)=tcool
	  temold(ix,iy)=tcool
	end do

	do ix=1,nx
	  do iy=1,ny
	    do j=1,6
	      p(ix,iy,j)=0.0d0
	    end do
	    sou(ix,iy)=0.0d0
	  end do
	end do
	if ( lres ) then
	else
	  do ix=1,nx     ! these numbers are purely arbitrary
	    htc(ix,1)=5.0d3
	    htc(ix,2)=5.0d3
	  end do
	end if

	return
	end
	subroutine calccoeffs(nx,ny,iter)
	include 'mgimp.inc'

	integer item,item1,item2,itcool,itin,iter
	double precision cond,rho,dd1,dd2,dd3,dd4
	double precision cond1,cond2,cond3,cond4,cbar
	double precision dmax99,gtemp,gtemp1,gtemp2,gt1,gt2,hbar
	dimension cond(0:mx+1,0:my+1)        
	
c       find properties over domain
	do ix=0,nx+1
	do iy=0,ny+1
	  item=temold(ix,iy)
	  cond(ix,iy)=data(item,4)
	end do
	end do
	itcool=tcool
	itin=tinlet
	rho=data(itcool,10)  ! rho must be constant by mass continuity
	dd1=difcut*rho*speed
	dd2=rho*speed
	dd3=rho*speed*glat

c  dd4 is coeff used for inlet boundary condition
	dd4=rho*speed*data(itin,7)

C
c  interior points first
	do ix=2,nx-1
	  do iy=2,ny-1
	    cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	    cond1=2./cond1
	    cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	    cond2=2./cond2
	    cond3=1./cond(ix+1,iy)+1./cond(ix,iy)
	    cond3=2./cond3
	    cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	    cond4=2./cond4
	    item=(temold(ix-1,iy)+temold(ix,iy))/2.
	    cbar=data(item,7)
	    p(ix,iy,1)=dmax99(0.0d0,cond1/dxg(ix)-dd1*cbar)+dd2*cbar
	    p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	    p(ix,iy,3)=dmax99(0.0d0,cond3/dxg(ix+1)-dd1*cbar)
	    p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	    p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	    p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1                 p(ix,iy,4)+p(ix,iy,6)
	    gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	    gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	    item1=temold(ix,iy)
	    item2=temold(ix-1,iy)
	    gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)


	    gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	    gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
c           if ( ix.eq.55 .and. iy.ge.7 .and. iy.le.9 ) then
c       write(46,998)iter,iy,data(temold(ix,iy),6),
c     1            data(temold(ix-1,iy),6),gtemp2,
c     1     temold(ix,iy),temold(ix-1,iy)
c       end if
	  end do
	end do
c 998   format(2x,i3,2x,i3,2x,f10.2,2x,f10.2,2x,f10.2,2x,f16.11,2x,f16.11)

c now boundaries
c start with inlet bottom corner
	ix=1
	iy=1
	cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	cond1=2./cond1
	cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	cond2=2./cond2
	cond3=1./cond(ix+1,iy)+1./cond(ix,iy)
	cond3=2./cond3
	cond4=cond(ix,iy)
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
	p(ix,iy,1)=cond1/dxu(ix)*2.+dd2*cbar
c       p(ix,iy,1)=dd4+dd2*cbar
	p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	p(ix,iy,3)=dmax99(0.0d0,cond3/dxg(ix+1)-dd1*cbar)
	if ( lq ) then
	  p(ix,iy,4)=0.0
	else
	  hbar=htc(ix,1)/(1.+htc(ix,1)*dyv(iy)/2./cond4)
	  p(ix,iy,4)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	if ( lq ) then
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
     1              - qflux(ix,1)*dxu(ix)/dyv(iy)
	else
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end if

c  along bottom boundary
	iy=1
	do ix=2,nx-1
	  cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	  cond1=2./cond1
	  cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	  cond2=2./cond2
	  cond3=1./cond(ix+1,iy)+1./cond(ix,iy)
	  cond3=2./cond3
	  cond4=cond(ix,iy)
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
	  p(ix,iy,1)=dmax99(0.0d0,cond1/dxg(ix)-dd1*cbar)+dd2*cbar
	  p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	  p(ix,iy,3)=dmax99(0.0d0,cond3/dxg(ix+1)-dd1*cbar)
	  write(6,*) lq,ix,iy,p(ix,iy,4)
	  if ( lq ) then
	    p(ix,iy,4)=0.0
	  else
	    hbar=htc(ix,1)/(1.d0+htc(ix,1)*dyv(iy)/2.d0/cond4)
c            write(41,*)ix,htc(ix,1),hbar
	    p(ix,iy,4)=hbar*dxu(ix)/dyv(iy)
	  end if
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  if ( lq ) then
	    gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	    gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	    item1=temold(ix,iy)
	    item2=temold(ix-1,iy)
	    gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	    gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	    gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
     1                 - qflux(ix,1)*dxu(ix)/dyv(iy)
c       write(48,*)lq,ix,qflux(ix,1),temold(ix-1,iy)
	  else
	    gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	    gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	    item1=temold(ix,iy)
	    item2=temold(ix-1,iy)
	    gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	    gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	    gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	  end if
	end do

c  exit bottom corner
	iy=1
	ix=nx
	cond1=1.d0/cond(ix-1,iy)+1.d0/cond(ix,iy)
	cond1=2.d0/cond1
	cond2=1.d0/cond(ix,iy)+1.d0/cond(ix,iy+1)
	cond2=2.d0/cond2
	cond4=cond(ix,iy)
	item=(temold(ix-1,iy)+temold(ix,iy))/2.d0
	cbar=data(item,7)
	p(ix,iy,1)=dmax99(0.0d0,cond1/dxg(ix)-dd1*cbar)+dd2*cbar
	p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	p(ix,iy,3)=0.0d0
	if ( lq ) then
	  p(ix,iy,4)=0.0d0
	else
	  hbar=htc(ix,1)/(1.d0+htc(ix,1)*dyv(iy)/2.d0/cond4)
	  p(ix,iy,4)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	if ( lq ) then
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
     1             - qflux(ix,1)*dxu(ix)/dyv(iy)
	else
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end if

c  exit plane
	ix=nx
	do iy=2,ny-1
	  cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	  cond1=2./cond1
	  cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	  cond2=2./cond2
	  cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	  cond4=2./cond4
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
	  p(ix,iy,1)=dmax99(0.0d0,cond1/dxg(ix)-dd1*cbar)+dd2*cbar
	  p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	  p(ix,iy,3)=0.0d0
	  p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end do

c  top exit corner
	ix=nx
	iy=ny
	cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	cond1=2./cond1
	cond2=cond(ix,iy)
	cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	cond4=2./cond4
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
	p(ix,iy,1)=dmax99(0.0d0,cond1/dxg(ix)-dd1*cbar)+dd2*cbar
	if ( lq ) then
	  p(ix,iy,2)=0.0d0
	else
	  hbar=htc(ix,2)/(1.+htc(ix,2)*dyv(iy)/2./cond2)
	  p(ix,iy,2)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,3)=0.0d0
	p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	item1=temold(ix,iy)
	item2=temold(ix-1,iy)
	gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	gtemp2= -dd2*( gt1-gt2 )
	sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	if ( lq ) then
	  sou(ix,iy)=sou(ix,iy)-qflux(ix,2)*dxu(ix)/dyv(iy)
	end if

c  top boundary
	iy=ny
	do ix=2,nx-1
	  cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	  cond1=2./cond1
	  cond2=cond(ix,iy)
	  cond3=1./cond(ix+1,iy)+1./cond(ix,iy)
	  cond3=2./cond3
	  cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	  cond4=2./cond4
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
	  p(ix,iy,1)=dmax99(0.0d0,cond1/dxg(ix)-dd1*cbar)+dd2*cbar
	  if ( lq ) then
	    p(ix,iy,2)=0.0d0
	  else
	    hbar=htc(ix,2)/(1.+htc(ix,2)*dyv(iy)/2./cond2)
	    p(ix,iy,2)=hbar*dxu(ix)/dyv(iy)
	  end if
	  p(ix,iy,3)=dmax99(0.0d0,cond3/dxg(ix+1)-dd1*cbar)
	  p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	  if ( lq ) then
	    sou(ix,iy)=sou(ix,iy)-qflux(ix,2)*dxu(ix)/dyv(iy)
	  end if
c         write(51,*)iter,ix,iy,sou(ix,iy),qflux(ix,2)
	end do

c  entry top corner
	ix=1
	iy=ny
	cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	cond1=2./cond1
	cond2=cond(ix,iy)
	cond3=1./cond(ix+1,iy)+1./cond(ix,iy)
	cond3=2./cond3
	cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	cond4=2./cond4
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
	p(ix,iy,1)=cond1/dxu(ix)*2.+dd2*cbar
c       p(ix,iy,1)=dd4+dd2*cbar
	if ( lq ) then
	  p(ix,iy,2)=0.0d0
	else
	  hbar=htc(ix,2)/(1.+htc(ix,2)*dyv(iy)/2./cond2)
	  p(ix,iy,2)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,3)=dmax99(0.0d0,cond3/dxg(ix+1)-dd1*cbar)
	p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	item1=temold(ix,iy)
	item2=temold(ix-1,iy)
	gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	gtemp2= -dd2*( gt1-gt2 )
	sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	if ( lq ) then
	  sou(ix,iy)=sou(ix,iy)-qflux(ix,2)*dxu(ix)/dyv(iy)
	end if

c  rest of inlet boundary
	ix=1
	do iy=2,ny-1
	  cond1=1./cond(ix-1,iy)+1./cond(ix,iy)
	  cond1=2./cond1
	  cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	  cond2=2./cond2
	  cond3=1./cond(ix+1,iy)+1./cond(ix,iy)
	  cond3=2./cond3
	  cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	  cond4=2./cond4
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
	  p(ix,iy,1)=cond1/dxu(ix)*2.+dd2*cbar
c         p(ix,iy,1)=dd4+dd2*cbar
	  p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	  p(ix,iy,3)=dmax99(0.0d0,cond3/dxg(ix+1)-dd1*cbar)
	  p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end do

c       goto 231
c       do ix=1,nx
c       do iy=1,ny
c       sou(ix,iy)=0.
c       p(ix,iy,5)=p(ix,iy,5)-p(ix,iy,6)
c       end do
c       end do
c231    continue
	return
	end
	subroutine calccoeffsp(nx,ny,iter)
	include 'mgimp.inc'

       
	integer item,item1,item2,itcool,itin,iter
	double precision cond,rho,dd1,dd2,dd3,dd4
	double precision cond1,cond2,cond3,cond4,cbar
	double precision dmax99,gtemp,gtemp1,gtemp2,gt1,gt2,hbar
	dimension cond(0:mx+1,0:my+1)

c       find properties over domain
	do ix=0,nx+1
	do iy=0,ny+1
	  item=temold(ix,iy)
	  cond(ix,iy)=data(item,4)
	end do
	end do

	itcool=tcool
	itin=tinlet
	rho=data(itcool,10)  ! rho must be constant by mass continuity
	dd1=difcut*rho*speed
	dd2=rho*speed
	dd3=rho*speed*glat

c  dd4 is coeff used for inlet boundary condition
	dd4=rho*speed*data(itin,7)

        write(6,1998)data(item+1,6)
1998    format(f16.11)

c  interior points first
	do ix=2,nx-1
	  do iy=2,ny-1
	    cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	    cond2=2./cond2
	    cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	    cond4=2./cond4
	    item=(temold(ix-1,iy)+temold(ix,iy))/2.
	    cbar=data(item,7)
	    p(ix,iy,1)=dd2*cbar
	    p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	    p(ix,iy,3)=0.0d0
	    p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	    p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	    p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1                 p(ix,iy,4)+p(ix,iy,6)
	    gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	    gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	    item1=temold(ix,iy)
	    item2=temold(ix-1,iy)
	    gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	    gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	    gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
c           if ( ix.eq.55 .and. iy.ge.7 .and. iy.le.9 ) then
c       write(46,998)iter,iy,data(temold(ix,iy),6),
c     1            data(temold(ix-1,iy),6),gtemp2,
c     1     temold(ix,iy),temold(ix-1,iy)
c       end if
	  end do
	end do
c 998   format(2x,i3,2x,i3,2x,f10.2,2x,f10.2,2x,f10.2,2x,f16.11,2x,f16.11)

c now boundaries
c start with inlet bottom corner
	ix=1
	iy=1
	cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	cond2=2./cond2
	cond4=cond(ix,iy)
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
c        cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	p(ix,iy,1)=dd2*cbar
c       p(ix,iy,1)=dd4+dd2*cbar
	p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	p(ix,iy,3)=0.0d0
	if ( lq ) then
	  p(ix,iy,4)=0.0d0
	else
	  hbar=htc(ix,1)/(1.+htc(ix,1)*dyv(iy)/2./cond4)
	  p(ix,iy,4)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
c       write(33,*)ix,htc(ix,1),hbar,p(ix,iy,2),p(ix,iy,4),p(ix,iy,5)
	if ( lq ) then
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
     1              - qflux(ix,1)*dxu(ix)/dyv(iy)
	else
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
c       write(37,*)ix,p(ix,iy,6)*temold(ix,iy),sou(ix,iy),p(ix,iy,2)
c     1     ,p(ix,iy,4)
	end if

c  along bottom boundary
	iy=1
	do ix=2,nx-1
	  cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	  cond2=2./cond2
	  cond4=cond(ix,iy)
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
c          cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	  p(ix,iy,1)=dd2*cbar
	  p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	  p(ix,iy,3)=0.0d0
	  if ( lq ) then
	    p(ix,iy,4)=0.0d0
	  else
	    hbar=htc(ix,1)/(1.+htc(ix,1)*dyv(iy)/2./cond4)
	    p(ix,iy,4)=hbar*dxu(ix)/dyv(iy)
	  end if
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  if ( lq ) then
	    gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	    gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	    item1=temold(ix,iy)
	    item2=temold(ix-1,iy)
	    gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	    gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	    gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
     1                 - qflux(ix,1)*dxu(ix)/dyv(iy)
c       write(48,*)lq,ix,qflux(ix,1),temold(ix-1,iy)
	  else
	    gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	    gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	    item1=temold(ix,iy)
	    item2=temold(ix-1,iy)
	    gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	    gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	    gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	  end if
	end do

c  exit bottom corner
	iy=1
	ix=nx
	cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	cond2=2./cond2
	cond4=cond(ix,iy)
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
c        cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	p(ix,iy,1)=dd2*cbar
	p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	p(ix,iy,3)=0.0d0
	if ( lq ) then
	  p(ix,iy,4)=0.0d0
	else
	  hbar=htc(ix,1)/(1.+htc(ix,1)*dyv(iy)/2./cond4)
	  p(ix,iy,4)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	if ( lq ) then
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
     1             - qflux(ix,1)*dxu(ix)/dyv(iy)
	else
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end if

c  exit plane
	ix=nx
	do iy=2,ny-1
	  cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	  cond2=2./cond2
	  cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	  cond4=2./cond4
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
c          cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	  p(ix,iy,1)=dd2*cbar
	  p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	  p(ix,iy,3)=0.0d0
	  p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end do

c  top exit corner
	ix=nx
	iy=ny
	cond2=cond(ix,iy)
	cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	cond4=2./cond4
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
c        cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	p(ix,iy,1)=dd2*cbar
	if ( lq ) then
	  p(ix,iy,2)=0.0d0
	else
	  hbar=htc(ix,2)/(1.+htc(ix,2)*dyv(iy)/2./cond2)
	  p(ix,iy,2)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,3)=0.0d0
	p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	item1=temold(ix,iy)
	item2=temold(ix-1,iy)
	gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	gtemp2= -dd2*( gt1-gt2 )
	sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	if ( lq ) then
	  sou(ix,iy)=sou(ix,iy)-qflux(ix,2)*dxu(ix)/dyv(iy)
	end if

c  top boundary
	iy=ny
	do ix=2,nx-1
	  cond2=cond(ix,iy)
	  cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	  cond4=2./cond4
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
c          cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	  p(ix,iy,1)=dd2*cbar
	  if ( lq ) then
	    p(ix,iy,2)=0.0d0
	  else
	    hbar=htc(ix,2)/(1.+htc(ix,2)*dyv(iy)/2./cond2)
	    p(ix,iy,2)=hbar*dxu(ix)/dyv(iy)
	  end if
	  p(ix,iy,3)=0.0d0
	  p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	  sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	  if ( lq ) then
	    sou(ix,iy)=sou(ix,iy)-qflux(ix,2)*dxu(ix)/dyv(iy)
	  end if
c         write(51,*)iter,ix,iy,sou(ix,iy),qflux(ix,2)
	end do

c  entry top corner
	ix=1
	iy=ny
	cond2=cond(ix,iy)
	cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	cond4=2./cond4
	item=(temold(ix-1,iy)+temold(ix,iy))/2.
	cbar=data(item,7)
c        cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	p(ix,iy,1)=dd2*cbar
c       p(ix,iy,1)=dd4+dd2*cbar
	if ( lq ) then
	  p(ix,iy,2)=0.0d0
	else
	  hbar=htc(ix,2)/(1.+htc(ix,2)*dyv(iy)/2./cond2)
	  p(ix,iy,2)=hbar*dxu(ix)/dyv(iy)
	end if
	p(ix,iy,3)=0.0d0
	p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1             p(ix,iy,4)+p(ix,iy,6)
	gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
c       write(31,*)ix,htc(ix,2),hbar,p(ix,iy,2),p(ix,iy,4)
c     1  ,p(ix,iy,5)
	gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	item1=temold(ix,iy)
	item2=temold(ix-1,iy)
	gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	gtemp2= -dd2*( gt1-gt2 )
	sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	if ( lq ) then
	  sou(ix,iy)=sou(ix,iy)-qflux(ix,2)*dxu(ix)/dyv(iy)
	end if

c  rest of inlet boundary
	ix=1
	do iy=2,ny-1
	  cond2=1./cond(ix,iy)+1./cond(ix,iy+1)
	  cond2=2./cond2
	  cond4=1./cond(ix,iy)+1./cond(ix,iy-1)
	  cond4=2./cond4
	  item=(temold(ix-1,iy)+temold(ix,iy))/2.
	  cbar=data(item,7)
c          cbar=data((temold(ix-1,iy)+temold(ix,iy))/2.,7)
	  p(ix,iy,1)=dd2*cbar
c         p(ix,iy,1)=dd4+dd2*cbar
	  p(ix,iy,2)=cond2/dyg(iy+1)*dxu(ix)/dyv(iy)
	  p(ix,iy,3)=0.0d0
	  p(ix,iy,4)=cond4/dyg(iy)*dxu(ix)/dyv(iy)
	  p(ix,iy,6)=rho*cbar*dxu(ix)/dtfals
	  p(ix,iy,5)=p(ix,iy,1)+p(ix,iy,2)+p(ix,iy,3)+
     1               p(ix,iy,4)+p(ix,iy,6)
	  gtemp=dd3*(frold(ix,iy)-frold(ix-1,iy))
	  gtemp1=dd2*cbar*(temold(ix,iy)-temold(ix-1,iy))
	  item1=temold(ix,iy)
	  item2=temold(ix-1,iy)
	  gt1=data(item1,6)+data(item1,7)*(temold(ix,iy)-item1)
	  gt2=data(item2,6)+data(item2,7)*(temold(ix-1,iy)-item2)
	  gtemp2= -dd2*( gt1-gt2 )
	    sou(ix,iy)=gtemp+gtemp1+gtemp2+p(ix,iy,6)*temold(ix,iy)
	end do

	return
	end

	subroutine solve(nx,ny)
c  solves iteratively a set of linear difference equations
c  pcoe is the matrix of coefficients and sou are the sources
c  tem is the array of values to be found
c  the solver is 2-dimensional and assumes that the two directions involved
c  are nominally called x and y .
c  the coefficient array is dimensioned as p(nx,ny,5) where nx is the
c  number of cells in the x direction and ny the number of cells in the y
c  direction . the dimension 5 refers to the adjoining cells as follows .
c
c                       2
c
c                  |----------|
c                  |          |
c           1      |    5     |    3
c                  |          |         y ( j )
c                  |          |         ^
c                  |----------|         |
c                                       |
c                       4                ---> x ( i )
c
c
c  the equation to be solved is of the form
c
c   p(i,j,5)*tem(i,j)=sum { p(i,j,k)*tem(ip,jp) } + sou(i,j)
c                       k=1,4             ( ip jp refer to neighbours)
c
c   here p(i,j,5)=sum { p(i,j,k) } [ unless the source contributes ]
c                   k=1,4
c         ( + false time step term in this case )
c
c
c  the array tem has been dimensioned to 0:nx+1,0:ny+1 so that the sources
c  can be either explicity prescribed and the coefficients on the boundary
c  set to zero or the extra values can represent external values.
c
c
c  for 3-d cases the other direction can be put in explicity by defining
c  the source appropiately i.e. in terms of the coefficient 5 and the source.
c  the third direction values not changing each iteration (slab by slab method )
c
c
c
	include 'mgimp.inc'


	dimension gn(0:mx+1,0:my+1),ge(0:mx+1,0:my+1),gb(0:mx+1,0:my+1)
	integer iter
	double precision gn,ge,gb,deno,gto

	do i=1,nx+1
	do j=0,ny+1
	 gn(i,j)=0.0d0
	 ge(i,j)=0.0d0
	 gb(i,j)=0.0d0
	end do
	end do
	i=0
	do j=0,ny+1
	 gn(i,j)=0.0d0
	 ge(i,j)=0.0d0
	 gb(i,j)=tem(i,j)
	end do
	j=0
	do i=0,nx+1
	 gn(i,j)=0.0d0
	 ge(i,j)=0.0d0
	 gb(i,j)=tem(i,j)
	end do
	do iter=1,liter
c   ***** the four following compiler directives seem to give the
c         same answers ( with -O2 option , but not -O3 option ) as
c         unoptimised version ( halves compute time ).
c         however leave out for now as i don't fully trust it.
c   ***** it is a recursive algorithm
cc  c$dir no_recurrence
	  do i=1,nx
cc  c$dir no_recurrence
	  do j=1,ny
	    deno=p(i,j,5)-p(i,j,4)*gn(i,j-1)-p(i,j,1)*ge(i-1,j)
	    gn(i,j)=p(i,j,2)/deno
	    ge(i,j)=p(i,j,3)/deno
	    gto=p(i,j,4)*ge(i,j-1)*tem(i+1,j-1)
     1         +p(i,j,1)*gn(i-1,j)*tem(i-1,j+1)
     1         +p(i,j,4)*gb(i,j-1)+p(i,j,1)*gb(i-1,j)+sou(i,j)
	    gb(i,j)=gto/deno
	  end do
	  end do
cc  c$dir no_recurrence
	  do j=ny,1,-1
cc  c$dir no_recurrence
	  do i=nx,1,-1
	    tem(i,j)=gn(i,j)*tem(i,j+1)+ge(i,j)*tem(i+1,j)+gb(i,j)
	  end do
	  end do
	end do
	return
	end


	subroutine solvep(nx,ny)
c  solves iteratively a set of linear difference equations
c  pcoe is the matrix of coefficients and sou are the sources
c  tem is the array of values to be found
c  the solver is 2-dimensional and assumes that the two directions involved
c  are nominally called x and y .
c  solver assumes parabolic in positive x
c  the coefficient array is dimensioned as p(nx,ny,5) where nx is the
c  number of cells in the x direction and ny the number of cells in the y
c  direction . the dimension 5 refers to the adjoining cells as follows .
c
c                       2
c
c                  |----------|
c                  |          |
c           1      |    5     |    3
c                  |          |         y ( j )
c                  |          |         ^
c                  |----------|         |
c                                       |
c                       4                ---> x ( i )
c
c
c  the equation to be solved is of the form
c
c   p(i,j,5)*tem(i,j)=sum { p(i,j,k)*tem(ip,jp) } + sou(i,j)
c                       k=1,4             ( ip jp refer to neighbours)
c
c   here p(i,j,5)=sum { p(i,j,k) } [ unless the source contributes ]
c                   k=1,4
c         ( + false time step term in this case )
c
c
c  the array tem has been dimensioned to 0:nx+1,0:ny+1 so that the sources
c  can be either explicity prescribed and the coefficients on the boundary
c  set to zero or the extra values can represent external values.
c
c
c
c  parabolic so no conduction in x i.e. p(i,j,3)=0.0
c
c
	include 'mgimp.inc'


	dimension en(my),fn(my)
	double precision en,fn,p1,p2,p3,p4,p5,si
	do j=1,ny
	  en(j)=0.0d0
	  fn(j)=0.0d0
	end do
	do ix=1,nx
	  p2=p(ix,1,2)
	  en(2)=p(ix,1,5)/p2
	  fn(2)=-(sou(ix,1)+p(ix,1,1)*tem(ix-1,1)
     1            +p(ix,1,4)*tem(ix,0) )/p2
	  do iy=3,ny
	    p4=p(ix,iy-1,4)
	    p2=p(ix,iy-1,2)
	    p5=p(ix,iy-1,5)
	    p1=p(ix,iy-1,1)
	    si=sou(ix,iy-1)+p1*tem(ix-1,iy-1)
	    en(iy)=(p5*en(iy-1)-p4)/
     1             (p2*en(iy-1))
	    fn(iy)=(p5*fn(iy-1)-p2*en(iy)*fn(iy-1)-si)/p2
	  end do
	  p4=p(ix,ny,4)
	  p2=p(ix,ny,2)
	  p5=p(ix,ny,5)
	  p1=p(ix,ny,1)
	  si=sou(ix,ny)+p1*tem(ix-1,ny)+p2*tem(ix,ny+1)
	  tem(ix,ny)=(en(ny)*si-p4*fn(ny))/
     1               (en(ny)*p5-p4)
	  do iy=ny,2,-1
	    tem(ix,iy-1)=(tem(ix,iy)-fn(iy))/en(iy)
	  end do
	end do
	return
	end

	subroutine fluxint(nx)

	include 'mgimp.inc'
	integer nz,ni,i1,i2,k,ni4
	double precision z1,dz,fsum1,fsum2,z2,f0,f1,f2,f3,f4,f5
	double precision fsum,gzintp
C-----------------------------------------------------------------------------

	nz=nx
	ni=nflux
	ni4=ni
	ni=ni+2
	xc(ni-1)=xc(ni-2)+1.e-3
	gmeasflux(ni-1,1)=0.0d0
	gmeasflux(ni-1,2)=0.0d0
	xc(ni)=1.0d2
	gmeasflux(ni,1)=0.0d0
	gmeasflux(ni,2)=0.0d0
c
	i1=1
	i2=2
	open(89,file='hf.dat',status='unknown')
	write(89,*)' I X QBOT QTOP '
	do 30 j=1,nz
	 if (j.ne.1) then
	    z1=xu(j-1)
	    dz=xu(j)-xu(j-1)
	  else
	    z1=0.0d0
	    dz=xu(j)
	  endif
	  dz=dz*0.1
	  fsum=0.0d0
	  fsum1=0.0d0
	  do 40 k=1,10
	    z2=z1+dz
	    f1=gzintp(z1,ni,i1)
	    f2=gzintp(z2,ni,i1)
	    f3=gzintp(z1,ni,i2)
	    f4=gzintp(z2,ni,i2)
	    f0=(f1+f2)*0.5
	    f5=(f3+f4)*0.5
	    z1=z2
	    fsum=fsum+f0
40          fsum1=fsum1+f5
	    qflux(j,1)=fsum*0.1
	    qflux(j,2)=fsum1*0.1
	    write(89,991)j,xu(j),qflux(j,1),qflux(j,2)
30      continue
	close(89)
	return
991     format(2x,i3,2x,f5.3,2x,2(f10.2,2x))
	end

	real*8 function gzintp(z0,ni,i0)
	include 'mgimp.inc'
	integer ni,i0,jm1
	double precision tiny,gdif,z0

	j=1
	tiny=-1.e-6
	gdif=z0-xc(1)
	if (gdif.le.tiny ) then
	  write(14,*) 'mistake when interpolating heat flux...z=',z0
	  stop
	endif

10      j=j+1
	if (j.gt.ni) goto 50
	if (z0.lt.xc(j)) goto 20
	goto 10
20      jm1=j-1
	if ( abs(xc(j)-xc(jm1)) .lt. abs(tiny) ) goto 60
	gzintp=gmeasflux(jm1,i0)+(gmeasflux(j,i0)-
     1         gmeasflux(jm1,i0))*
     1         (z0-xc(jm1))/(xc(j)-xc(jm1))
	goto 100
50      write(14,*) 'insufficient heat flux data'
	write(14,*) 'ran out before end of domain'
	write(14,*) 'z0, xc(ni), j and ni are ',z0,xc(ni),j,ni
	stop
60      write(14,*) 'mistake in interpolation....'
	write(14,*) 'two points in data refer to same distance'
	stop
100     return
	end

	subroutine update(nx,ny)
	include 'mgimp.inc'
	double precision cre1,cre2


c put new into old

	cre1=1.-relaxh
	cre2=1.-relaxf
	do iy=1,ny
	  do ix=1,nx
	    temold(ix,iy)=relaxh*tem(ix,iy)+
     1                    cre1*temold(ix,iy)
	    frold(ix,iy)=relaxf*fr(ix,iy)+
     1                    cre2*frold(ix,iy)
	  end do
	end do
	return
	end
	subroutine calconv(nx,ny,iconv,iter,resref)
	include 'mgimp.inc'
	integer ixm,iym,iconv,iter
	double precision sum1,gmax,gmax1,sum,sum2,resid,resref


	sum1=0.0d0
	gmax=0.0d0
	gmax1=0.0d0
	ixm=0
	iym=0
	do iy=1,ny
	  do ix=1,nx
	    sum=p(ix,iy,5)*tem(ix,iy)
	    sum=sum-p(ix,iy,1)*tem(ix-1,iy)
	    sum=sum-p(ix,iy,2)*tem(ix,iy+1)
	    sum=sum-p(ix,iy,3)*tem(ix+1,iy)
	    sum=sum-p(ix,iy,4)*tem(ix,iy-1)
	    sum=sum-sou(ix,iy)
	    sum2=abs(sum)
	    sum1=sum1+sum2
	    if ( sum2.gt.gmax ) then
	      ixm=ix
	      iym=iy
	      gmax=sum2
	      gmax1=sum
	    end if
	  end do
	end do
	resid=sum1/resref
	if ( resid .le. gfrac ) iconv=1
	write(6,999)iter,resid
c       write(6,*)
c       write(6,998)gmax1,ixm,iym
c       write(6,*)
	return
c998    format(' greatest residual is ',f11.1,' at i= ',i3,'
c     1 and j= ',i3)
999    format(' summed relative residual at iter ',i4,' is ',f10.4)
	end
	real*8 function dmin99(a,b)
	implicit none
	double precision a,b
	if ( a . ge . b ) then
	 dmin99 = b
	else
	 dmin99 = a
	end if
	return
	end
	integer*4 function jdint99(a)
	implicit none
	double precision a
	jdint99=a
	return
	end
	real*8 function dmax99(a,b)
	implicit none
	double precision a,b
	if ( a.le.b ) then
	  dmax99=b
	else
	  dmax99=a
	end if
	return
	end
	subroutine updatespeed(avt,nx,ny,isp,ites3)
	include 'mgimp.inc'
	integer isp,ites3,ites4
	double precision avt,gtes3

c find av exit temp

	avt=0.0d0
	do iy=1,ny
	  avt=avt+tem(nx,iy)*dyv(iy)
	end do
	avt=avt/yv(ny)
	gtes3=avt-texit
	if ( gtes3 .gt. 0 ) then
	  ites4=1
	else
	  ites4=-1
	end if
	if ( ites3.eq.ites4 ) then
	else
	  gtes4=gtes4/2.
	end if
	if ( abs(gtes3) .gt. gclose ) then ! gclose defines closeness
c                                       ! of answer
	  speed=speed*(1.-gtes3*gtes4/abs(gtes3))
	  isp=0
	else
	  isp=1
	end if
	ites3=ites4
	return
	end
	subroutine dumpsol(nx,ny)
	include 'mgimp.inc'
	integer it,jdint99
	double precision frt


	open(47,file='RES.DAT',status='unknown')
	write(47,*)nx,ny
	do ix=1,nx
	  do iy=1,ny
	    write(47,898)xg(ix),yg(iy),tem(ix,iy),fr(ix,iy)
	  end do
	end do
898     format(1x,f7.4,2x,f7.4,2x,f9.2,2x,f7.4)
	close(47)

c restart file
	open(47,file='RESTART.DAT',status='unknown')
	write(47,*)nx,ny
	do iy=1,ny
	  write(47,*)(tem(ix,iy),ix=1,nx)
	  write(47,*)(fr(ix,iy),ix=1,nx)
	end do
	close(47)
	open(47,file='RESC.DAT',status='unknown')
	do ix=1,nx
	  do iy=1,ny
	    it=jdint99(tem(ix,iy))
	    frt=data(it,2)+(tem(ix,iy)-dfloat(it))*data(it,3)
	    write(47,*)fr(ix,iy),frt
	  end do
	end do
	close(47)
	if ( lq ) then
	  open(50,file='HFLUX.DAT',status='unknown')
	  write(50,50)
	  do 491 ix=1,nx
	    write(50,51) ix,xg(ix),htc(ix,1),qflux(ix,1),
     1          htc(ix,2),qflux(ix,2),fsint(ix),airgap(ix)
491        continue
	else
	  open(50,file='HTC.DAT',status='unknown')
	  write(50,50)
	  do 49 ix=1,nx
	    write(50,51) ix,xg(ix),htc(ix,1),qflux(ix,1),
     1          htc(ix,2),qflux(ix,2),fsint(ix),airgap(ix)
49        continue
	end if
50      format(2x,1hN,4x,1hx,11x,4hHTOP,8x,4hFLXT,8x,
     1            4hHBOT,8x,4hFLXB,8x,4hFINT,8x,4hAIRG)
51      format(2x,I3,2x,E10.4,2x,E10.4,2x,E10.4,2x,
     1            E10.4,2x,E10.4,2x,E10.4,2x,E10.4)
	close(50)
	return
	end
	subroutine updatef(nx,ny,ihint)
	include 'mgimp.inc'
	integer ifix,k,jfix,jdint99,ihint
	double precision h,hs,hl,dh,grlx,grlxm1,dec,cpt,frl
	double precision frnew,add,frrlx,dmin99,dmax99

c---   calculate enthalpies for each cell
	hs=enth(0,1)
	hl=enth(ihint,1)
	dh=1.d0/((hl-hs)*1.d-3)
	grlx=relaxf
	grlxm1=1.d0-grlx
	do 10 j=1,ny
	 do 10 i=1,nx
	  jfix=jdint99(tem(i,j))
	  dec=tem(i,j)-dfloat(jfix)
c---   calculate sensible enthalpy:
	  cpt=data(jfix,6)+dec*data(jfix,7)
c---   calculate latent enthalpy:
	  frl=(1.d0-fr(i,j))*glat
c---   total enthalpy:
	  h=cpt+frl
       if (h.le.hs) then
	frnew=1.d0
       else if (h.ge.hl) then
	frnew=0.0d0
       else
	k=jdint99((h-hs)*dh)
	add=h-enth(k,1)
	frnew=enth(k,4)+add*enth(k,5)*enth(k,3)
       end if
c---   relax fraction solid
	frrlx=fr(i,j)*grlxm1+grlx*frnew
	fr(i,j)=dmin99(1.d0,dmax99(0.d0,frrlx))
c        if ( i.eq.nx .and. j.eq.ny ) then
c        write(6,*)relaxf,frrlx,frnew,h,hs,hl,tem(i,j)
c        end if
   10  continue
c       write(6,*)fr(5,ny),fr(100,25)
       return
c 994    format(2x,f6.2,2x,f8.1,2x,f8.1,2x,f9.7,2x,f9.7)
       end



	subroutine calcaux(nx,ny)
	include 'mgimp.inc'

	dimension work1(mx),work2(mx)
	integer item
	double precision shrnkl,dthick,work1,work2,dmax99,dmin99
	double precision tiny,fsintp,cond4,hbar

c---  convert volume shrinkage to a linear shrinkage:
	shrnkl=shrnk*0.33333d0
c---  calculate thickness solidified (in work1) for each x-posn down caster
c---  & reduction in gap due to taper (in work2)
	 dthick=dyv(1)
	do 10 ix=1,nx
	 work1(ix)=dthick*dmax99(fr(ix,1)-fstart,0.0d0)
	 if ( xu(ix) .le. gx ) then
	  work2(ix)=xu(ix)/gx*taper
	 else
	  work2(ix)=taper
	 end if
10      continue
	 do 20 iy=2,ny
	  dthick=dyv(iy)
	   do 20 ix=1,nx
	     work1(ix)=work1(ix)+dthick*
     1                 dmax99(fr(ix,iy)-fstart,0.0d0)
20       continue
c---  calculate airgap and mean fs across thickness
	 do 30 ix=1,nx
	  fsint(ix)=work1(ix)/gy
	  airgap(ix)=work1(ix)*shrnkl-work2(ix)
30       continue
c-------------------------------------------------------------------
c---
c---  calculate htc based on fsint values:
c---
c---   if fsint < fmax:
c---   htc = hmax
c---
c---  if fmax < fsint < fsqueeze:
c---   htc = hpar1 + hpar2*( hpar3*fsint +
c---                           hpar4*(fsint^hpar5))
c---  if fsqueeze < fsint:
c---   htc = hpar1 + hpar2*( hpar3*fsqueeze +
c---                           hpar4*(fsint^hpar5))
c---   hmax >= htc >= hmin
c---
c------------------------------------------------------------------
c

	if ( lq ) then
	  tiny=1.e-6
	  do 35 ix=1,nx
	    htc(ix,1)=qflux(ix,1)/(tem(ix,1)-tcool+tiny)
	    htc(ix,2)=qflux(ix,2)/(tem(ix,ny)-tcool+tiny)
35        continue
	else
	  do 40 ix=1,nx
	    fsintp=fsint(ix)
	    if ( fsintp .lt. fmax ) then
	      htc(ix,1)=hmax
	    else if (fsintp.lt.fsqueeze) then
	      htc(ix,1)=hpar1+
     1                  hpar2*(hpar3*fsintp+hpar4*(fsintp**hpar5))
	    else
	      htc(ix,1)=hpar1+
     1                  hpar2*(hpar3*fsqueeze+hpar4*(fsintp**hpar5))
	    endif
c---
c---   add-in squeezing component due to taper:
	    if (airgap(ix).lt.0.0d0) then
	      htc(ix,1)=htc(ix,1)-airgap(ix)*hpar6
	      airgap(ix)=0.0d0
	    endif
c---
	    htc(ix,1)=dmax99(hmin,dmin99(htc(ix,1),hmax))
	    if ( xu(ix) .gt. gx ) htc(ix,1)=0.0d0
	    htc(ix,2)=htc(ix,1)
	    item=tem(ix,1)
	    cond4=data(item,4)
	    hbar=htc(ix,1)/(1.+htc(ix,1)*dyv(1)/2./cond4)
	    qflux(ix,1)=hbar*(tem(ix,1)-tcool)
	    qflux(ix,2)=hbar*(tem(ix,ny)-tcool)
c       write(97,*)ix,htc(ix,1),tem(ix,1)
40        continue
	end if
c--
c--
	return
	end
