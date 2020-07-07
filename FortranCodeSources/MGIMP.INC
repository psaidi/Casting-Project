        implicit none
        integer mx,my,nd
        parameter ( mx= 500 , my = 500 ,nd = 1200 )
c        parameter ( mx= 200 , my = 50 ,nd = 1200 )
        integer nx,ny,i,j,ix,iy,nflux,ialloy,nsweep,liter,niter
        logical lq,ls,lres,lsol
        character*20 file1
        double precision gx,gy,xu,xg,dxu,dxg,yv,yg,dyv,dyg
        double precision p,sou,tem,fr,temold,frold,qflux,htc
        double precision shrink,xc,gmeasflux,fsint,airgap
        double precision data,enth,glat
        double precision tinlet,texit,speed,tcool,taper
        double precision difcut,relaxh,relaxf,dtfals,gtes4,gfrac
        double precision gclose,shrnk,fstart,fsqueeze,
     1                   hmax,fmax,hmin,
     1                   hpar1,hpar2,hpar3,
     1                   hpar4,hpar5,hpar6
        common /geom/ gx,gy
c  length of caster in x ( along caster ) and y ( through thickness )

        common /gridx/ xu(0:mx),xg(mx),dxu(mx),dxg(mx)
        common /gridy/ yv(0:my),yg(my),dyv(my),dyg(my)
c  grid geom in standard phoenics type speak

        common /coeff/ p(mx,my,6),sou(mx,my)
c  coefficients of equations and source terms

        common /var/ tem(0:mx+1,0:my+1),fr(0:mx,my),
     1               temold(0:mx+1,0:my+1),frold(0:mx,my)
c  temperature and fraction solid distribution

        common /surf/ qflux(mx,2),htc(mx,2),
     1             shrink(mx,2),xc(mx),gmeasflux(mx,2),
     1             fsint(mx),airgap(mx),nflux
c  heat flux, h.t.c. and shrinkage distributions on surface
c  plus measured heat flux data in x and gmeasflux
c  plus integrated fraction solid and airgap
        common /props/ data(0:nd,11),enth(0:nd,5),
     1              glat,ialloy
c  general property array, props in mushy zone latent heat and alloy number

        common /proc/ tinlet,texit,speed,tcool,taper,nsweep
c  inlet temp, exit temp and exit speed

        common /logic/ lq,ls,lsol
c  lq=t ==> heat flux data used for calcs ( otherwise h.t.c's calculated )
c  ls=t ==> speed is updated ( given exit temperature )
c  lsol=t ==> solver is whole field; else parabolic

        common /soln/ difcut,relaxh,relaxf,dtfals,gtes4,gfrac,gclose,
     1                liter,niter
c  solution parameters

        common /htcparam/ shrnk,fstart,fsqueeze,
     1                    hmax,fmax,hmin,
     1                    hpar1,hpar2,hpar3,
     1                    hpar4,hpar5,hpar6
c parameters needed for dynamic gap calc ( see subroutine calcaux )


