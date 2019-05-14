 'Declare a new WeibullDataSet object. 
  Dim WDS As New WeibullDataSet
 
 'Specify the analysis settings. 
  WDS.AnalysisSettings.Analysis = WeibullSolverMethod_RRX 'RRX(0), RRY(1), MLE(2) 
  'HAVE TO SELECT 2 MLE IF USING BAYESIAN WEIBULL-> for Bayesian only
  'add a check on the data to see if >10 points then MLE, if last/largest hours point is suspension --> MLE
  WDS.AnalysisSettings.ConfBounds = WeibullSolverCBMethod_FisherMatrix '0=FM, 1 = Likelihood ratio, 2 = betaNinomial for Mixed weibull only, HAVE TO SELECT 3 IF USING BAYESIAN WEIBULL-> for Bayesian only
  WDS.AnalysisSettings.Distribution = WeibullSolverDistribution_Weibull '0 = Weibull, 1 = Normal, 3 = Lognormal, 9 = Bayesian
  WDS.AnalysisSettings.Parameters = WeibullSolverNumParameters_MS_2Parameter '0 1-parameter, 1 = 2p, 3 = 3p, 3 = MS_MIXED (only for Weibull distribution)
  WDS.AnalysisSettings.RankingMethod = WeibullSolverRankMethod_Median '0 = median ranking method, 1 = Kalan-Meier estimator
  
  'other parameters
  WDS.AnalysisSettings.SortBeforeCalculations = True
  WDS.AnalysisSettings.UngroupGroupedData = False
  WDS.AnalysisSettings.UseRSRegression = False
  WDS.AnalysisSettings.UseSpecialSort = True
  WDS.AnalysisSettings.Weibull_UnbiasParameters = False

 'Add failure times to the data set. 
  Call WDS.AddFailure(100, 1)
  Call WDS.AddFailure(120, 1)
  call WDS.AddSuspension(150,1)

  '_____________________________________________________________________________________
  'BEST FIT - DISTRIBUTION WIZARD ANALYSIS
  '_____________________________________________________________________________________
  'Consider the normal, lognormal and 2-parameter Weibull distributions in the evaluation. 
  WDS.BestFitSettings.AllowExponential1 = False
  WDS.BestFitSettings.AllowExponential2 = False
  WDS.BestFitSettings.AllowNormal = True
  WDS.BestFitSettings.AllowLognormal = True
  WDS.BestFitSettings.AllowWeibull2 = True
  WDS.BestFitSettings.AllowWeibull3 = False
  WDS.BestFitSettings.AllowGamma = False
  WDS.BestFitSettings.AllowGenGamma = False
  WDS.BestFitSettings.AllowLogistic = False
  WDS.BestFitSettings.AllowLoglogistic = False
  WDS.BestFitSettings.AllowGumbel = False
 
 'Use the MLE parameter estimation method. 
  WDS.BestFitSettings.Analysis = WeibullSolverMethod_MLE
  
 'Determine which distribution best fits the data set, based on the MLE method. 
  Call WDS.CalculateBestFit()
  '_____________________________________________________________________________________

 'Analyze the data set. 
  'WDS.Calculate()
  
  'Calculate the reliability at 100 hrs and display the result. 
  Dim r As Double
  r = WDS.FittedModel.Reliability(100)
  MsgBox("Reliability at 100 hrs: " & r)
  
  '________________________________________
  'Bayesian Prior Distributions
  '________________________________________
  'beta std deviation and mean from prior data set
  
  
  'Bayesian specific set-up
    WDS.AnalysisSettings.Weibull_Bayesian_PriorDistribution = 0 '0 sets to B-W normal distribution (1 is lognormal, 2 is exponential, 3 is uniform)
	WDS.AnalysisSettings.Weibull_Bayesian_mean = 2 'value of the mean of beta from prior data set
	WDS.AnalysisSettings.Weibull_Bayesian_Std = 0.4 'value of the standard deviation of beta from prior data set
	WDS.AnalysisSettings.Weibull_Bayesian_ResultsAs = 1 '1 means it uses the mean of beta, 0 means it used median
	