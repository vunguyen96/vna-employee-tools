"use strict";(self.webpackChunkvna_employee_tools=self.webpackChunkvna_employee_tools||[]).push([[590],{3590:(X,u,n)=>{n.r(u),n.d(u,{MailModule:()=>J});var t=n(4438),R=n(177),p=n(604),f=n(9192),r=n(855),g=n(5911),C=n(8834),D=n(728),v=n(9631),m=n(9417),h=n(2102),b=n(9213),x=n(7554),F=n(1413),y=n(5964),A=n(6977);let E=(()=>{class e{constructor(a,o,i){this.msalGuardConfig=a,this.authService=o,this.msalBroadcastService=i,this.isIframe=!1,this.loginDisplay=!1,this._destroying$=new F.B}ngOnInit(){this.authService.instance.initialize().then(()=>{this.authService.handleRedirectObservable().subscribe(),this.isIframe=window!==window.parent&&!window.opener,this.authService.instance.enableAccountStorageEvents(),this.msalBroadcastService.inProgress$.pipe((0,y.p)(a=>a===x.T$.None),(0,A.Q)(this._destroying$)).subscribe(()=>{this.checkAndSetActiveAccount()})})}checkAndSetActiveAccount(){if(!this.authService.instance.getActiveAccount()&&this.authService.instance.getAllAccounts().length>0){let o=this.authService.instance.getAllAccounts();this.authService.instance.setActiveAccount(o[0])}}loginRedirect(){this.msalGuardConfig.authRequest?this.authService.loginRedirect({...this.msalGuardConfig.authRequest}):this.authService.loginRedirect()}ngOnDestroy(){this._destroying$.next(void 0),this._destroying$.complete()}static#t=this.\u0275fac=function(o){return new(o||e)(t.rXU(r.LR),t.rXU(r.dl),t.rXU(r.qb))};static#e=this.\u0275cmp=t.VBU({type:e,selectors:[["app-ms-login"]],decls:4,vars:0,consts:[[3,"click"]],template:function(o,i){1&o&&(t.j41(0,"p"),t.EFF(1," ms-login works! "),t.j41(2,"button",0),t.bIt("click",function(){return i.loginRedirect()}),t.EFF(3,"Login"),t.k0s()())}})}return e})();class S{constructor(){this.getDisplayAddress=()=>`${this.emailAddress.name?`${this.emailAddress.name} <${this.emailAddress.address.substring(0,10)}...>`:this.emailAddress.address.substring(0,10)}`}}var T=n(3352),M=n(5596),l=n(4069);function $(e,s){1&e&&(t.j41(0,"th",21),t.EFF(1," Sender "),t.k0s())}function w(e,s){if(1&e&&(t.j41(0,"td",22)(1,"p",23),t.EFF(2),t.k0s()()),2&e){const a=s.$implicit;t.R7$(2),t.JRh(a.sender)}}function I(e,s){1&e&&(t.j41(0,"th",21),t.EFF(1," Subject "),t.k0s())}function O(e,s){if(1&e&&(t.j41(0,"td",22)(1,"p",24),t.EFF(2),t.k0s()()),2&e){const a=s.$implicit;t.R7$(2),t.JRh(a.subject)}}function N(e,s){1&e&&(t.j41(0,"th",21),t.EFF(1," Sent Date Time "),t.k0s())}function k(e,s){if(1&e&&(t.j41(0,"td",22)(1,"p"),t.EFF(2),t.k0s()()),2&e){const a=s.$implicit;t.R7$(2),t.JRh(a.sentDateTime)}}function G(e,s){1&e&&(t.j41(0,"th",21),t.EFF(1," To Recipients "),t.k0s())}function P(e,s){if(1&e&&(t.j41(0,"td",22)(1,"p",23),t.EFF(2),t.k0s()()),2&e){const a=s.$implicit;t.R7$(2),t.JRh(a.toRecipients)}}function B(e,s){1&e&&t.nrm(0,"tr",25)}function H(e,s){1&e&&t.nrm(0,"tr",26)}function V(e,s){1&e&&(t.j41(0,"tr",27)(1,"td",28),t.EFF(2,"No available data"),t.k0s()())}let L=(()=>{class e{constructor(a){this.graphService=a,this.displayedColumns=["sender","subject","sentDateTime","toRecipients"],this.dataSource=[],this.mailSearchCriteria={subject:"Weekly",senderEmail:"vu.nguyen@ptnglobalcorp.com"}}ngOnInit(){}ngAfterContentInit(){}search(){this.graphService.getGraphClient().api("/me/mailFolders/SentItems/messages").filter(`sentDateTime ge 2024-01-01  and contains(Subject, '${this.mailSearchCriteria.subject}') and contains(from/emailAddress/address,'${this.mailSearchCriteria.senderEmail}')`).top(20).orderby(["sentDateTime desc"]).get().then(o=>{this.dataSource=o.value.map(i=>({sender:Object.assign(new S,i.sender).getDisplayAddress(),subject:i.subject,sentDateTime:new Date(i.sentDateTime).toDateString(),toRecipients:i.toRecipients.map(d=>Object.assign(new S,d).getDisplayAddress()).join(", ")})),console.log(this.dataSource)})}static#t=this.\u0275fac=function(o){return new(o||e)(t.rXU(T.G))};static#e=this.\u0275cmp=t.VBU({type:e,selectors:[["app-mailbox"]],decls:35,vars:5,consts:[[1,"card"],[1,"w-100","card-body","d-flex","justify-content-center","align-items-center"],[1,"w-100"],["appearance","outline",1,"fs-16","w-auto"],["matInput","","type","text","placeholder","Enter subject",3,"ngModelChange","ngModel"],["appearance","outline",1,"fs-16","w-auto","mx-3"],["matInput","","type","email","placeholder","Enter sender email",3,"ngModelChange","ngModel"],["mat-icon-button","",3,"click"],["aria-label","icon-button with a search icon"],[1,"mb-56"],[1,"table-responsive"],["mat-table","",1,"text-nowrap","w-100",3,"dataSource"],["matColumnDef","sender"],["mat-header-cell","",4,"matHeaderCellDef"],["mat-cell","",4,"matCellDef"],["matColumnDef","subject"],["matColumnDef","sentDateTime"],["matColumnDef","toRecipients"],["mat-header-row","",4,"matHeaderRowDef"],["mat-row","",4,"matRowDef","matRowDefColumns"],["class","mat-row",4,"matNoDataRow"],["mat-header-cell",""],["mat-cell",""],[1,"mb-0","fw-medium"],[1,"mb-0","fw-medium","op-5"],["mat-header-row",""],["mat-row",""],[1,"mat-row"],["colspan","4",1,"mat-cell"]],template:function(o,i){1&o&&(t.j41(0,"div",0)(1,"mat-card",1)(2,"mat-card-content",2)(3,"mat-form-field",3)(4,"mat-label"),t.EFF(5,"Subject"),t.k0s(),t.j41(6,"input",4),t.mxI("ngModelChange",function(c){return t.DH7(i.mailSearchCriteria.subject,c)||(i.mailSearchCriteria.subject=c),c}),t.k0s()(),t.j41(7,"mat-form-field",5)(8,"mat-label"),t.EFF(9,"Sender Email"),t.k0s(),t.j41(10,"input",6),t.mxI("ngModelChange",function(c){return t.DH7(i.mailSearchCriteria.senderEmail,c)||(i.mailSearchCriteria.senderEmail=c),c}),t.k0s()(),t.j41(11,"button",7),t.bIt("click",function(){return i.search()}),t.j41(12,"mat-icon",8),t.EFF(13,"search"),t.k0s()()()()(),t.j41(14,"mat-card",2)(15,"mat-card-content")(16,"h4",9),t.EFF(17,"Inbox"),t.k0s(),t.j41(18,"div",10)(19,"table",11),t.qex(20,12),t.DNE(21,$,2,0,"th",13)(22,w,3,1,"td",14),t.bVm(),t.qex(23,15),t.DNE(24,I,2,0,"th",13)(25,O,3,1,"td",14),t.bVm(),t.qex(26,16),t.DNE(27,N,2,0,"th",13)(28,k,3,1,"td",14),t.bVm(),t.qex(29,17),t.DNE(30,G,2,0,"th",13)(31,P,3,1,"td",14),t.bVm(),t.DNE(32,B,1,0,"tr",18)(33,H,1,0,"tr",19)(34,V,3,0,"tr",20),t.k0s()()()()),2&o&&(t.R7$(6),t.R50("ngModel",i.mailSearchCriteria.subject),t.R7$(4),t.R50("ngModel",i.mailSearchCriteria.senderEmail),t.R7$(9),t.Y8G("dataSource",i.dataSource),t.R7$(13),t.Y8G("matHeaderRowDef",i.displayedColumns),t.R7$(),t.Y8G("matRowDefColumns",i.displayedColumns))},dependencies:[v.fg,h.rl,h.nJ,m.me,m.BC,m.vS,b.An,C.iY,M.RN,M.m2,l.Zl,l.tL,l.ji,l.cC,l.YV,l.iL,l.KS,l.$R,l.YZ,l.NB,l.ky],styles:[".table-responsive[_ngcontent-%COMP%]{overflow-x:auto}td.mat-cell[_ngcontent-%COMP%]{padding:16px}th.mat-header-cell[_ngcontent-%COMP%]:first-of-type, td.mat-cell[_ngcontent-%COMP%]:first-of-type, td.mat-footer-cell[_ngcontent-%COMP%]:first-of-type{padding:16px}th.mat-header-cell[_ngcontent-%COMP%]{padding:16px}.mat-header-cell[_ngcontent-%COMP%]{font-size:14px}"]})}return e})();var U=n(4791),Y=n(4160);const j=[{path:"",component:p.H,children:[{path:"mailbox",component:L,canActivate:[r.VA]},{path:"ms-login",component:E}]}];let J=(()=>{class e{static#t=this.\u0275fac=function(o){return new(o||e)};static#e=this.\u0275mod=t.$C({type:e,bootstrap:[p.H,r.vI]});static#n=this.\u0275inj=t.G2t({providers:[(0,f.lh)(j),(0,t.oKB)(C.Hl,g.s5),(0,D.pg)()],imports:[R.MD,f.iI.forChild(j),g.s5,v.fS,m.YN,m.X1,h.RG,b.m_,U.E,Y.J]})}return e})()}}]);