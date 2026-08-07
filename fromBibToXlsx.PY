import pandas as pd
import re

# ------------------------------------------------------------
# 1. Load the Excel file
# ------------------------------------------------------------
file_path = "elsevier_citation.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# ------------------------------------------------------------
# 2. Define keyword sets (case‑insensitive)
# ------------------------------------------------------------
ai_terms = [
    "artificial intelligence", "machine learning", "deep learning",
    "neural network", "ai-driven", "intelligent systems",
    "support vector machine", "random forest", "gradient boosting",
    "xgboost", "lstm", "gru", "autoencoder", "generative adversarial",
    "isolation forest", "one-class", "ensemble learning", "transfer learning",
    "federated learning", "reinforcement learning", "bayesian network"
]

fraud_terms = [
    "fraud detection", "fraud prevention", "anomaly detection",
    "fraudulent transaction", "financial fraud", "credit card fraud",
    "payment fraud", "transaction fraud", "fraudulent activity",
    "anti-fraud", "fraud risk"
]

card_terms = [
    "smart card", "chip card", "emv", "contactless payment",
    "payment card", "debit card", "credit card", "pos transaction",
    "credit card fraud", "card transaction", "card payment",
    "prepaid card", "stored-value card"
]

# Pre‑compile regex patterns (word boundaries)
ai_re = re.compile(r'\b(?:' + '|'.join(re.escape(term) for term in ai_terms) + r')\b', re.IGNORECASE)
fraud_re = re.compile(r'\b(?:' + '|'.join(re.escape(term) for term in fraud_terms) + r')\b', re.IGNORECASE)
card_re = re.compile(r'\b(?:' + '|'.join(re.escape(term) for term in card_terms) + r')\b', re.IGNORECASE)

# ------------------------------------------------------------
# 3. Manual overrides from previous analysis (ID -> decision, reason)
#    (only for the entries already examined)
# ------------------------------------------------------------
manual_decisions = {
    1222: ("Exclude", "Focuses on case-based reasoning and explanation patterns (AI), but no mention of fraud detection or payment cards."),
    1223: ("Exclude", "Discusses KADS knowledge engineering (AI), but not fraud detection or payment cards."),
    589:  ("Exclude", "Uses rule induction (ML) for financial statement errors/fraud, but no payment card terms."),
    1212: ("Exclude", "Domain knowledge for database discovery (AI), but no fraud detection or payment card focus."),
    526:  ("Exclude", "Security agents for information retrieval (AI), but no fraud detection or payment card terms."),
    1087: ("Exclude", "Knowledge acquisition for rule-based expert systems (AI), but no fraud or payment card terms."),
    1207: ("Exclude", "Hybrid intelligent system (AI) for predicting bank holding structures – not fraud detection, no payment cards."),
    1209: ("Exclude", "General intelligent data analysis (AI), not specific to fraud or payment cards."),
    1094: ("Exclude", "Unexpectedness in knowledge discovery (AI) – mentions data mining, but no fraud or payment card terms."),
    1068: ("Exclude", "Survey of credit and behavioural scoring – AI/ML not mentioned, not fraud detection per se."),
    757:  ("Borderline", "Business data mining and ML – mentions fraud detection in passing, payment card terms appear only in a survey context."),
    1097: ("Exclude", "Knowledge management and data mining for marketing – no fraud detection, no payment card terms."),
    1210: ("Exclude", "Dynamic rule refinement in data mining – no fraud detection, no payment card terms."),
    1235: ("Exclude", "Expert systems for knowledge management – no fraud detection, no payment card terms."),
    1099: ("Include", "Explicitly discusses data mining techniques (including neural networks) for credit card fraud detection."),
    1219: ("Exclude", "Mining data value conversion rules – no fraud detection, no payment card terms."),
    640:  ("Exclude", "SVM ensemble – mentions fraud detection as an example but not the focus, no payment card terms."),
    704:  ("Exclude", "Statistical anomaly detection for httpd data – not financial fraud, no payment cards."),
    1213: ("Exclude", "Mining class outliers – mentions CRM applications but not specifically payment card fraud."),
    286:  ("Include", "Uses CART and MARS (ML) for credit scoring on bank credit card data. Fraud detection implied."),
    658:  ("Exclude", "Telecommunications subscription fraud – not payment cards."),
    1037: ("Exclude", "Mobile phone inspection process – not fraud detection, no payment cards."),
    1083: ("Exclude", "Rule extraction from neural networks – mentions data mining but no fraud or payment card focus."),
    1086: ("Exclude", "Telecom churn management – not payment cards."),
    1092: ("Exclude", "Cybercrime review – no specific payment card focus."),
    1095: ("Exclude", "Internet traffic anomaly detection – not financial fraud, no payment cards."),
    1214: ("Exclude", "Cellular phone test process – not fraud, not payment cards."),
    33:   ("Exclude", "Mobile user substitution detection – not payment cards."),
    1043: ("Exclude", "Cost-sensitive boosting for imbalanced data – general ML, no payment card focus."),
    1048: ("Exclude", "Kernel ellipsoidal trimming – general outlier detection, no payment card focus."),
    1:    ("Include", "Incorporates domain knowledge into data mining classifiers for indirect bank lending (credit cards). Fraud detection central."),
    215:  ("Include", "Fraud detection for plastic card transactions (credit/debit cards) using one-class and two-class classification."),
    1289: ("Exclude", "Classifier ensembles survey – mentions fraud as an example but no payment card focus."),
    245:  ("Include", "Survey of signature-based methods for financial fraud detection – explicitly covers credit card fraud."),
    510:  ("Exclude", "General data mining challenges – no specific fraud or payment card focus."),
    559:  ("Exclude", "Cluster-based under-sampling for imbalanced data – general, no fraud or payment card focus."),
    597:  ("Exclude", "Outlier identification for CRM – no payment card focus."),
    802:  ("Exclude", "Expert system for telecom fraud – not payment cards."),
    1000: ("Include", "Data mining for credit card default prediction – payment card term present, ML for risk/fraud detection."),
    1052: ("Exclude", "SVM for customer-centric data mining – no explicit fraud detection or payment card focus."),
    543:  ("Exclude", "Induction over strategic agents – no fraud or payment card focus."),
    584:  ("Exclude", "Knowledge extraction from MCLP – no fraud or payment card focus."),
    719:  ("Exclude", "Machine learning for radioxenon event classification – not financial fraud, no payment cards."),
    832:  ("Exclude", "Artificial immune system for online break-in fraud – not payment card specific."),
    1005: ("Include", "Subagging for credit scoring models – payment card (credit card) mentioned, ML used for risk prediction."),
    1011: ("Exclude", "Unsupervised learning for imbalanced datasets – no fraud or payment card focus."),
    1065: ("Exclude", "Density-based outlier mining – general, no payment card focus."),
    1077: ("Exclude", "Time series prediction – not fraud, no payment cards."),
    1202: ("Exclude", "Network visualization for enterprise networks – not fraud, no payment cards."),
    1243: ("Include", "Multiple classifiers for credit risk assessment – payment card (credit) implied, ML used. Risk assessment relevant."),
    1318: ("Exclude", "Selecting data mining for business decisions – no fraud or payment card focus."),
    1326: ("Exclude", "Online user identification – no fraud or payment card focus."),
    16:   ("Exclude", "Masquerade mimicry attack detection – computer security, not payment card fraud."),
    92:   ("Exclude", "Telecom subscription fraud – not payment cards."),
    251:  ("Exclude", "Anomaly detection with SOM and PSO – general, no payment card focus."),
    261:  ("Exclude", "Hybrid outlier detection – general, mentions credit card fraud but not the focus."),
    414:  ("Exclude", "Fraud in online games/lotteries – not payment cards."),
    476:  ("Exclude", "VoIP fraud survey – not payment cards."),
    547:  ("Exclude", "Financial statement fraud using linguistic analysis – no payment card terms."),
    555:  ("Exclude", "Fast outlier detection for log data – general, no payment card focus."),
    670:  ("Exclude", "Class noise detection for software metrics – not fraud, no payment cards."),
    672:  ("Exclude", "Financial statement fraud detection – no payment card terms."),
    694:  ("Exclude", "Probabilistic risk analysis for entry control – not payment cards."),
    864:  ("Exclude", "Online auction fraud detection – not payment card specific (though credit cards may be used)."),
    938:  ("Include", "Compares SVM and random forests for credit card fraud detection on real transaction data."),
    1017: ("Include", "Multiple classifier architectures for credit risk assessment – payment card implied."),
    1028: ("Exclude", "Intrusion detection using neural networks – not financial fraud, no payment cards."),
    1050: ("Exclude", "Support vector regression acceleration – general ML, no fraud/payment card focus."),
    1073: ("Exclude", "Random forests survey – mentions outlier detection but no payment card focus."),
    1100: ("Exclude", "SVM ensembles for imbalanced datasets – general, no payment card focus."),
    1115: ("Include", "Literature review of data mining for financial fraud detection – explicitly covers credit card fraud."),
    1253: ("Exclude", "Spam detection using AIS and neural networks – not payment card fraud."),
    1272: ("Exclude", "Software defect prediction – not fraud, no payment cards."),
    1306: ("Include", "Data mining for credit scoring (credit cards) – payment card term present, ML used for risk/fraud detection."),
    1343: ("Exclude", "Financial statement fraud detection – no payment card terms."),
    102:  ("Include", "Hierarchical linear SVM for credit card fraud detection – explicit mention."),
    114:  ("Exclude", "Health insurance fraud – not payment cards."),
    221:  ("Exclude", "Network anomaly detection – not financial fraud, no payment cards."),
    514:  ("Exclude", "Density-based outlier detection – general, no payment card focus."),
    585:  ("Exclude", "Outlier detection in data streams – general, no payment card focus."),
    590:  ("Exclude", "Dataset shift in classification – general, no fraud/payment card focus."),
    646:  ("Exclude", "Health insurance claim abuse detection – not payment cards."),
    843:  ("Exclude", "Stock fraud detection – not payment cards."),
    865:  ("Exclude", "Improved competitive learning for network intrusion/fraud detection – not payment card specific."),
    880:  ("Include", "Artificial immune classifier for credit scoring analysis – payment card (credit) mentioned."),
    1001: ("Exclude", "Preprocessing methods for class imbalance – general, no payment card focus."),
    1035: ("Exclude", "Global knowledge extraction – no fraud or payment card focus."),
    1374: ("Exclude", "Rule-based policy language for proactive fraud management – not payment card specific."),
    2:    ("Exclude", "Gravitation-based classification – general, no fraud/payment card focus."),
    28:   ("Exclude", "Botnet detection – not financial fraud, no payment cards."),
    78:   ("Exclude", "Real-time risk assessment on freeways – not fraud, no payment cards."),
    504:  ("Exclude", "Online neural networks for non-stationary data – general, no payment card focus."),
    565:  ("Include", "Feature selection for corporate credit rating prediction – payment card (credit) implied."),
    831:  ("Exclude", "Phishing detection – not payment card specific."),
    954:  ("Include", "Cost-sensitive decision tree for fraud detection on credit card data."),
    1012: ("Exclude", "Botnets survey – not financial fraud, no payment cards."),
    1014: ("Exclude", "Failure prediction from log files – not fraud, no payment cards."),
    1041: ("Exclude", "Linear discriminant analysis with concept drift – general, no payment card focus."),
    1225: ("Exclude", "VoIP spam detection – not payment cards."),
    1267: ("Exclude", "Outlier detection for patient monitoring – not financial fraud, no payment cards."),
    1286: ("Exclude", "Feature selection for MCLP – general, no payment card focus."),
    1304: ("Exclude", "ADCs for secure application delivery – not fraud detection."),
    1385: ("Exclude", "Energy usage modeling in smart homes – not fraud, no payment cards."),
    55:   ("Exclude", "Multiple classifier systems survey – general, mentions fraud but no payment card focus."),
    77:   ("Exclude", "Benford's law for Internet traffic – not fraud, no payment cards."),
    85:   ("Exclude", "Feature selection using DEA – general, no payment card focus."),
    390:  ("Include", "Credit card fraud detection – covers imbalanced data, non-stationarity, assessment."),
    424:  ("Exclude", "Online auction shilling fraud – not payment card specific."),
    505:  ("Exclude", "Big Data survey – general, no fraud or payment card focus."),
    548:  ("Exclude", "Multi-class decision-theoretic rough sets – general, no payment card focus."),
    572:  ("Exclude", "Natural gas theft detection – not payment cards."),
    592:  ("Exclude", "DNS security – not fraud detection."),
    1007: ("Exclude", "Clustering with mixture models – general, no payment card focus."),
    1032: ("Exclude", "Kernel entropy manifold learning for financial data – no explicit fraud detection, no payment card terms."),
    1055: ("Exclude", "Novelty detection for bank failure prediction – not fraud detection per se, no payment card terms."),
    1231: ("Exclude", "Robust outlier detection – general, mentions finance but no payment card focus."),
    1366: ("Exclude", "Mobile malware detection – not payment card fraud."),
    22:   ("Exclude", "Big Data in railway assets – not fraud, no payment cards."),
    74:   ("Exclude", "Mobile phone spam filtering – not payment cards."),
    84:   ("Exclude", "Enterprise purchase prediction – not fraud, no payment cards."),
    255:  ("Exclude", "Network intrusion detection – not financial fraud, no payment cards."),
    275:  ("Exclude", "Healthcare insurance fraud – not payment cards."),
    364:  ("Include", "APATE approach for credit card transaction fraud detection."),
    501:  ("Exclude", "Intrusion detection using MCLP and PSO – not financial fraud, no payment cards."),
    506:  ("Exclude", "Financial statement fraud detection – no payment card terms."),
    558:  ("Include", "Clustered SVM for credit scoring – payment card (credit) mentioned."),
    580:  ("Exclude", "Big Data solutions in management – general, no fraud or payment card focus."),
    581:  ("Include", "Survival mixture models for behavioral scoring – payment card (credit) implied."),
    588:  ("Borderline", "Outlier detection in data streams – mentions credit transactions but not the focus."),
    691:  ("Exclude", "One-class learning – general, no payment card focus."),
    892:  ("Exclude", "Survey of anomaly detection – general, no payment card focus."),
    1040: ("Exclude", "Random Balance for imbalanced data – general, no payment card focus."),
    1081: ("Exclude", "Bankruptcy prediction using neural networks – not fraud detection, no payment cards."),
    1091: ("Exclude", "SDN security survey – not payment card fraud."),
    1096: ("Exclude", "Ensemble method for imbalanced data – general, no payment card focus."),
    1160: ("Include", "Example-dependent cost-sensitive decision trees for credit card fraud detection."),
    1204: ("Exclude", "Privacy-preserving data publishing with outliers – not fraud detection."),
    1217: ("Exclude", "kNN-based outlier detection – general, no payment card focus."),
    1220: ("Exclude", "Microblog user activity prediction – not fraud, no payment cards."),
    1251: ("Exclude", "User identity verification – not payment card fraud."),
    1307: ("Exclude", "Behavioral analytics (short article) – no details, not specific."),
    1369: ("Include", "Big data fraud risk management at Alibaba – mentions credit cards, ML models."),
    1387: ("Exclude", "Anti-fraud technology for customer experience – general, no specific ML or payment card focus."),
    12:   ("Exclude", "Network intrusion detection using AIS – not payment card fraud."),
    13:   ("Exclude", "Rogue behavior detection in NoSQL graphs – not payment card specific."),
    58:   ("Exclude", "Ensemble outlier detection – general, no payment card focus."),
    60:   ("Exclude", "Evolutionary computing in CRM survey – mentions fraud but not payment card specific."),
    61:   ("Exclude", "Cost-sensitive meta-learning – general, no payment card focus."),
    101:  ("Exclude", "Distributed density-based outlier detection – general, mentions credit card fraud but not the focus."),
    105:  ("Include", "Survey of anomaly detection in financial domain – covers credit card fraud."),
    117:  ("Exclude", "Non-parameter outlier detection – general, mentions credit card fraud but not the focus."),
    158:  ("Exclude", "Anomaly prediction in data streams – general, mentions credit card fraud but not the focus."),
    216:  ("Exclude", "Network-wide traffic anomaly detection – not financial fraud."),
    268:  ("Exclude", "Industrial control systems anomaly detection – not payment cards."),
    273:  ("Exclude", "Social network anomaly detection – not payment cards."),
    297:  ("Exclude", "Payment intention prediction in mail order – not specifically payment cards."),
    397:  ("Include", "Feature engineering for credit card fraud detection."),
    434:  ("Exclude", "E-commerce identity fraud detection – not specifically payment cards."),
    516:  ("Exclude", "Digital versus physical payment use in China – not fraud detection."),
    544:  ("Exclude", "Hierarchical data stream clustering – general, no fraud/payment card focus."),
    551:  ("Exclude", "Privacy leakage to PayPal – not fraud detection."),
    564:  ("Exclude", "Associative classification for phishing detection – not payment card specific."),
    582:  ("Exclude", "Fusion of non-independent detectors – general, no payment card focus."),
    657:  ("Exclude", "Late payment prediction models – not fraud detection."),
    796:  ("Exclude", "Construction injury prediction – not fraud, no payment cards."),
    819:  ("Exclude", "Auto insurance fraud detection – not payment cards (wrong domain)."),
    1018: ("Exclude", "Dimensionality reduction for mixed-type data – general, no fraud/payment card focus."),
    1044: ("Exclude", "Managerial work and data triad – not fraud detection."),
    1053: ("Exclude", "Classification performance measure – general, no payment card focus."),
    1058: ("Exclude", "Root cause identification of security alerts – not payment card specific."),
    1062: ("Exclude", "Data leakage prevention systems survey – not specifically payment card fraud."),
    1084: ("Exclude", "Shilling groups in e-commerce – not payment card specific."),
    1205: ("Exclude", "Big Data for supply chain management – not fraud detection."),
    1262: ("Exclude", "Financial misstatement detection – not payment card specific."),
    1265: ("Exclude", "Network intrusion detection – not financial fraud."),
    1274: ("Exclude", "Health insurance fraud detection – not payment cards."),
    1299: ("Exclude", "ECG biometric authentication – not fraud detection."),
    1360: ("Exclude", "Social big data survey – general, no payment card focus."),
    5:    ("Exclude", "Financial restatements detection – not payment card specific."),
    32:   ("Exclude", "Data mining in accounting review – mentions fraud but no payment card focus."),
    38:   ("Exclude", "Sensor network security – not payment card fraud."),
    83:   ("Exclude", "Disruptive trading behavior detection – not payment cards."),
    200:  ("Exclude", "Outlier cluster detection – general, mentions credit card fraud but not the focus."),
    240:  ("Exclude", "Machine learning algorithms on big data – general, no payment card focus."),
    246:  ("Exclude", "Spatial anomaly detection in sensor networks – not payment cards."),
    250:  ("Exclude", "Anomaly detection using one-class SVM for system calls – not payment cards."),
    327:  ("Include", "Credit card fraud detection in e-tail – direct match."),
    509:  ("Exclude", "Event handling in production – not fraud, no payment cards."),
    522:  ("Borderline", "Data transformation for credit behavioral scoring – payment card implied, but risk assessment."),
    532:  ("Exclude", "Virtual property theft detection – not payment cards."),
    545:  ("Exclude", "SOM-based partial labeling for imbalanced data streams – general, no payment card focus."),
    546:  ("Exclude", "Social network security survey – not payment card fraud."),
    553:  ("Exclude", "SVM with priority class – general, no payment card focus."),
    561:  ("Exclude", "Gas sensor classification – not payment cards."),
    566:  ("Exclude", "Adaptive ensemble for imbalanced data – general, no payment card focus."),
    577:  ("Exclude", "Multi-kernel feature selection for imbalanced data – general, no payment card focus."),
    583:  ("Exclude", "Fusion analytics for policy insights – general, no payment card focus."),
    606:  ("Exclude", "Annotation error detection in speech corpora – not fraud, no payment cards."),
    619:  ("Exclude", "Dimensionality reduction – general, no payment card focus."),
    638:  ("Exclude", "Knowledge graph for abnormal expenses – not payment card specific."),
    714:  ("Exclude", "Fraud transaction detection in mail-order industry – not specifically payment cards."),
    809:  ("Exclude", "Financial statement fraud detection – no payment card terms."),
    1016: ("Exclude", "Social relation key for spam/SMS detection – not payment cards."),
    1023: ("Exclude", "Real-time classification of data streams – general, no payment card focus."),
    1027: ("Exclude", "Social network analytics for churn prediction – not fraud detection."),
    1034: ("Exclude", "Distributed computational model – not fraud, no payment cards."),
    1039: ("Exclude", "Local outlier mining – general, no payment card focus."),
    1067: ("Exclude", "Opinion mining from app reviews – not fraud detection."),
    1089: ("Borderline", "Cost-sensitive boosted tree for peer-to-peer lending – not credit cards, but online lending."),
    1093: ("Exclude", "Complex event processing – general, no payment card focus."),
    1129: ("Exclude", "Deep learning for credit scoring using CDS – not payment card specific."),
    1201: ("Exclude", "Clustering algorithms – general, no fraud/payment card focus."),
    1275: ("Exclude", "Solar panel malfunction detection – not fraud, no payment cards."),
    1323: ("Exclude", "Fog computing for anti-theft – not payment card specific."),
    65:   ("Exclude", "Cost-sensitive SVM – general, no payment card focus."),
    70:   ("Exclude", "Secure communication for smart city – not fraud detection."),
    80:   ("Exclude", "SIMO oversampling for imbalanced data – general, no payment card focus."),
    139:  ("Exclude", "Weight-based outlier detection – general, mentions credit card fraud but not the focus."),
    148:  ("Exclude", "Internal fraud in project-based organization – not payment cards."),
    179:  ("Exclude", "Semi-supervised multi-kernel ELM – general, mentions credit card fraud but not the focus."),
    233:  ("Exclude", "Hybrid intrusion detection – not financial fraud, no payment cards."),
    235:  ("Exclude", "Fraud detection in mail-order industry – not specifically payment cards."),
    247:  ("Exclude", "AI in IMRT planning (medical) – not fraud, no payment cards."),
    277:  ("Exclude", "Big Data healthcare framework – not fraud detection."),
    323:  ("Include", "Fraud-BNC customized Bayesian classifier for credit card fraud detection."),
    332:  ("Exclude", "Anomaly detection in smart card logs (transportation) – smart card present, but for transit, not financial fraud."),
    346:  ("Include", "Credit card fraud detection using big data analytics and ML."),
    348:  ("Include", "Risk Induced Bayesian Inference Bagging for credit card fraud detection."),
    355:  ("Include", "Sequence classification for credit card fraud detection using LSTM."),
    399:  ("Include", "SCARFF scalable framework for credit card fraud detection with Spark."),
    412:  ("Exclude", "Hybrid anomaly detection framework – not payment card specific."),
    517:  ("Exclude", "Neuro-fuzzy classifier for data streams – general, no payment card focus."),
    550:  ("Exclude", "Big data techniques in auditing – mentions fraud but not payment card specific."),
    563:  ("Exclude", "Big data in healthcare policy – not fraud detection."),
    586:  ("Exclude", "k-means SMOTE for imbalanced learning – general, no payment card focus."),
    596:  ("Exclude", "Mobile security awareness taxonomy – not fraud detection."),
    607:  ("Exclude", "Big data analytics survey – general, no payment card focus."),
    617:  ("Exclude", "Combining anomaly detectors for software security – not payment cards."),
    660:  ("Exclude", "Anomaly ranking with mixed data – general, no payment card focus."),
    683:  ("Exclude", "Big data frameworks survey – general, no payment card focus."),
    751:  ("Borderline", "Visual analytics for fraud detection – covers banks, stock market, telecom, insurance – not specifically payment cards."),
    755:  ("Exclude", "Outlier detection algorithms survey – general, mentions fraud but no payment card focus."),
    808:  ("Exclude", "Anomaly detection for cloud services – not payment cards."),
    817:  ("Exclude", "Knowledge base ontology for fraud detection – general, no payment card focus."),
    846:  ("Exclude", "Anomaly detection in press-hardening (manufacturing) – not payment cards."),
    873:  ("Exclude", "Neural networks for business computing – general, no fraud/payment card focus."),
    883:  ("Exclude", "Text mining for service quality in banks – not fraud detection."),
    900:  ("Include", "Cost-sensitive payment card fraud detection using random forest and kNN."),
    951:  ("Exclude", "Deep learning with LDA for automobile insurance fraud – not payment cards."),
    993:  ("Include", "Sequential fraud detection for prepaid cards using HMM."),
    1009: ("Exclude", "Data mining for smart production – not fraud detection."),
    1045: ("Include", "Credit risk prediction using financial/non-financial variables – payment card (credit) implied."),
    1047: ("Exclude", "Multi-class imbalance learning – general, no payment card focus."),
    1057: ("Exclude", "Active learning for crisis management – not payment card fraud."),
    1076: ("Exclude", "Rule extraction for tram faults – not fraud, no payment cards."),
    1079: ("Exclude", "Process deviation detection in trauma resuscitation – not fraud, no payment cards."),
    1088: ("Exclude", "Overlap-sensitive margin classifier – general, no payment card focus."),
    1153: ("Exclude", "ConvNets for telecom fraud detection – not payment cards."),
    1155: ("Exclude", "GAN for telecom fraud detection at receiving bank – not payment cards (telecom fraud)."),
    1226: ("Exclude", "Asymmetric entropy for imbalanced data – general, no payment card focus."),
    1234: ("Exclude", "Card-not-present e-commerce architectures – not fraud detection methods."),
    1258: ("Exclude", "Risk factor significance in cardiovascular diseases – not fraud, no payment cards."),
    1395: ("Exclude", "Handling data irregularities in classification – general, no payment card focus."),
    1426: ("Exclude", "Network intrusion detection – not financial fraud."),
    3:    ("Exclude", "Cost-sensitive SVM – general, no payment card focus."),
    54:   ("Exclude", "Density peak clustering – general, no payment card focus."),
    59:   ("Exclude", "AI in retailing – general, no fraud/payment card focus."),
    81:   ("Exclude", "Outlier detection in web usage mining – not payment cards."),
    86:   ("Exclude", "Malware detection in private cloud – not payment cards."),
    94:   ("Exclude", "Dynamic clustering for evolving environments – general, no payment card focus."),
    98:   ("Exclude", "Cost-sensitive adaptive boosting – mentions insurance fraud but not payment cards specifically."),
    109:  ("Exclude", "LSTM anomaly detection in supply chain – not payment card fraud."),
    110:  ("Exclude", "Secure element for sensor anomaly detection – not payment cards."),
    131:  ("Include", "Survey of ML for financial fraud detection – covers credit card fraud."),
    150:  ("Exclude", "DNS data exfiltration detection – not payment card fraud."),
    197:  ("Exclude", "ML for wireless communications in IoT – general, no fraud/payment card focus."),
    206:  ("Exclude", "Consumer-facing technology fraud – covers Internet, mobile, telecom – not specifically payment cards."),
    228:  ("Exclude", "Unsupervised network anomaly detection – not financial fraud."),
    230:  ("Exclude", "Mobile malware attacks – not payment card fraud."),
    243:  ("Exclude", "Anomaly detection via adaptive greedy model – general, no payment card focus."),
    269:  ("Exclude", "GAN for cancer-staging data – not fraud, no payment cards."),
    326:  ("Include", "GAN for improving classification in credit card fraud detection."),
    353:  ("Include", "Credit card fraud detection strategies with consumer incentives."),
    354:  ("Include", "Credit card fraud detection using ML on streaming data."),
    356:  ("Include", "Champion-challenger analysis for credit card fraud detection (hybrid ensemble and deep learning)."),
    375:  ("Include", "Feature construction for fraudulent credit card cash-out detection."),
    402:  ("Exclude", "Bankruptcy forecasting – not fraud detection, no payment card focus."),
    410:  ("Exclude", "Fraud threat to biometrics – general, no ML or payment card focus."),
    455:  ("Exclude", "Java source code anomaly detection – not payment card fraud."),
    462:  ("Exclude", "Abnormal passenger out-flow detection – not fraud, no payment cards."),
    475:  ("Exclude", "Cyber incident detection for EMR systems – not payment card fraud."),
    507:  ("Exclude", "Tax evasion detection – not payment card fraud."),
    515:  ("Exclude", "Economic perspective of big data security – general, no fraud methods."),
    528:  ("Exclude", "Outlier detection with Gaussian process – general, no payment card focus."),
    535:  ("Exclude", "Under-sampling for imbalanced data – general, no payment card focus."),
    537:  ("Exclude", "Forecasting sales in supply chain – not fraud detection."),
    576:  ("Exclude", "Compromised user credentials detection – not payment card specific."),
    595:  ("Exclude", "Fintech literature review – mentions fraud but no specific methods or payment card focus."),
    610:  ("Exclude", "Big data system survey – general, no fraud/payment card focus."),
    642:  ("Exclude", "Data fusion in smart city applications – general, not fraud specific."),
    652:  ("Exclude", "Implementing big data strategies – general, no fraud methods."),
    661:  ("Exclude", "Energy efficiency prediction – not fraud, no payment cards."),
    687:  ("Exclude", "IoT security survey – not fraud detection."),
    702:  ("Exclude", "GAN for financial time-series – not fraud detection."),
    767:  ("Exclude", "Deep neural network IDS for cloud – not financial fraud."),
    769:  ("Borderline", "Credit default prediction on peer-to-peer lending – not credit cards, but online lending."),
    800:  ("Exclude", "Outlier detection with autoencoder ensemble – general, mentions fraud but no payment card focus."),
    822:  ("Exclude", "Big data for credit scoring using mobile phone data – not payment cards."),
    837:  ("Exclude", "Anomaly detection in home automation – not payment cards."),
    844:  ("Exclude", "GAN for anomaly detection in medical imaging – not payment cards."),
    856:  ("Exclude", "Anomaly detection in unmanned aerial vehicles – not payment cards."),
    878:  ("Exclude", "Double-level ensemble learning for anomaly detection – general, no payment card focus."),
    903:  ("Include", "Proactive fraud detection in credit card transactions using Fourier/Wavelet transforms."),
    913:  ("Include", "Multiple Consensus model for fraud detection in e-commerce transactions (credit cards)."),
    960:  ("Exclude", "Anomaly detection in wide area networks – not payment cards."),
    1004: ("Exclude", "Adaptive online extreme learning machine – general, no payment card focus."),
    1046: ("Exclude", "Internet of vehicles survey – not fraud detection."),
    1064: ("Exclude", "SMOTE analysis – general, no payment card focus."),
    1085: ("Exclude", "Active authentication using file system decoys – not payment card fraud."),
    1139: ("Exclude", "Anomaly detection in wireless sensor networks – not payment cards."),
    1154: ("Exclude", "Deep learning in cyber forensics – general, no payment card focus."),
    1203: ("Exclude", "Resampling ensemble for imbalanced data streams – general, no payment card focus."),
    1236: ("Exclude", "Predictive analytics paradigm – general, no fraud/payment card focus."),
    1242: ("Exclude", "Over-sampling algorithms – general, no payment card focus."),
    1244: ("Exclude", "Digital transformation in financial institutions – not fraud detection."),
    1263: ("Exclude", "Boosting ensemble for streaming data – general, no payment card focus."),
    1268: ("Exclude", "E-commerce bibliographic survey – general, not fraud specific."),
    1273: ("Exclude", "Testing distributed anomaly detectors – general, no payment card focus."),
    1281: ("Exclude", "Undersampling method – general, no payment card focus."),
    1293: ("Exclude", "Illegal construction waste dumping detection – not payment cards."),
    1324: ("Exclude", "Biometry for cognitive automotive systems – not fraud detection."),
    1332: ("Exclude", "Bagged SVM for emotion recognition – not fraud, no payment cards."),
    1338: ("Exclude", "Cost-sensitive meta-learning classifier – general, no payment card focus."),
    4:    ("Exclude", "Imbalanced data streams framework – general, no payment card focus."),
    23:   ("Exclude", "Federated learning review – general, no fraud/payment card focus."),
    30:   ("Exclude", "Localization for automatic vehicles – not fraud detection."),
    67:   ("Exclude", "Novelty detection in multi-class systems – general, no payment card focus."),
    68:   ("Exclude", "Oversampling with Laplacian eigenmaps – general, no payment card focus."),
    69:   ("Exclude", "Fused variable screening for imbalanced data – general, no payment card focus."),
    72:   ("Exclude", "Gamma distribution sampling for imbalanced data – general, no payment card focus."),
    88:   ("Include", "Graph databases for financial fraud detection – mentions credit card fraud."),
    95:   ("Exclude", "Outlier detection in rule bases – general, no payment card focus."),
    163:  ("Borderline", "Entropy-based proactive credit scoring – payment card implied, but risk assessment."),
    169:  ("Include", "Hybrid credit scoring model for credit fraud detection – payment card implied."),
    196:  ("Exclude", "BiGAN for anomaly detection – general, no payment card focus."),
    218:  ("Include", "Credit risk model comparison – uses ML, payment card (bank loan data, credit card)."),
    223:  ("Exclude", "Anomaly pattern recognition in nuclear fusion – not fraud, no payment cards."),
    241:  ("Exclude", "Mortality risk in cardiac surgery – not fraud, no payment cards."),
    266:  ("Exclude", "Data imbalance experimental evaluation – general, mentions fraud but no payment card focus."),
    295:  ("Exclude", "VAT audit case selection – not payment card fraud."),
    299:  ("Exclude", "Database intrusion detection – not payment card fraud."),
    307:  ("Include", "Credit card fraud detection using ensemble learning and pipelining."),
    308:  ("Include", "Feature selection with SVM for credit card risk identification."),
    310:  ("Include", "Weighted ELM for imbalanced classification applied to credit card fraud detection."),
    312:  ("Include", "Game theory approach for credit card fraud detection rule pool management."),
    317:  ("Include", "Autoencoder-based model for credit card fraud detection."),
    325:  ("Include", "Automated feature engineering for credit card fraud detection using HMMs."),
    350:  ("Include", "Cost-sensitive meta-learning ensemble for credit card fraud detection."),
    393:  ("Include", "SVM-RFE with hyperparameter optimization for credit card fraud detection."),
    404:  ("Exclude", "EM side-channel analysis for activity detection – not payment card fraud."),
    405:  ("Exclude", "Resampling for imbalanced data considering noise – general, no payment card focus."),
    430:  ("Include", "Combining predictions for peer-to-peer lending – not credit cards, but online lending."),
    436:  ("Exclude", "One-class SVM for IoT data – not payment card fraud."),
    441:  ("Exclude", "Noisy-sample removal for imbalanced classification – general, no payment card focus."),
    484:  ("Exclude", "Secure data analytics for smart grid – not payment cards."),
    503:  ("Exclude", "Stream processing architecture for IoT – not fraud detection."),
    512:  ("Exclude", "IoNT survey – not fraud detection."),
    513:  ("Exclude", "Security enhancement against insiders in manufacturing – not payment cards."),
    524:  ("Exclude", "Web application IDS – not financial fraud."),
    542:  ("Exclude", "Scalable KDE-based outlier detection – general, no payment card focus."),
    562:  ("Exclude", "Tourist walk outlier detection – general, no payment card focus."),
    569:  ("Exclude", "ITS security overview – not payment card fraud."),
    574:  ("Exclude", "Evolutionary approach to fraud management – insurance fraud, not payment cards."),
    587:  ("Exclude", "Anomaly pattern detection for streaming data – general, no payment card focus."),
    626:  ("Exclude", "Secure routing in IoT – not fraud detection."),
    651:  ("Exclude", "AI in commercial banks – mentions fraud but no specific methods or payment card focus."),
    655:  ("Exclude", "Network intrusion detection – not financial fraud."),
    666:  ("Exclude", "Payment behavior pattern recognition – not fraud detection."),
    677:  ("Borderline", "Decision support reliability in fraud detection – uses financial fraud detection task, but not specifically payment cards."),
    682:  ("Exclude", "Graph-theoretic phishing detection – not payment card specific."),
    703:  ("Exclude", "EM side-channel analysis for IoT forensics – not payment card fraud."),
    711:  ("Exclude", "Spam caller detection – not payment cards."),
    753:  ("Exclude", "Corruption prediction across countries – not payment card fraud."),
    765:  ("Exclude", "Fraudulent financial reporting and data analytics – not payment cards."),
    771:  ("Exclude", "Health status monitoring for diabetes – not fraud, no payment cards."),
    787:  ("Exclude", "Autoencoder for payment system anomaly detection – interbank payments, not consumer cards."),
    805:  ("Exclude", "Neural memory plasticity for medical anomaly detection – not payment cards."),
    885:  ("Include", "Intuitionistic fuzzy logic decision tree for e-transactional fraud detection – covers credit cards."),
    969:  ("Include", "Fraud detection in bank transaction using deep RNN – covers credit cards."),
    976:  ("Exclude", "Anomaly detection in transportation networks – not payment cards."),
    989:  ("Exclude", "Graph-based anomaly detection for fraud – systematic review, not specifically payment cards."),
    1019: ("Exclude", "Heterogeneous online learning ensemble – general, no payment card focus."),
    1024: ("Exclude", "Blockchain in hospitality – not fraud detection."),
    1025: ("Exclude", "Nonprofit fraud reporting – not payment cards."),
    1026: ("Exclude", "Fintech research directions – general, no specific methods."),
    1029: ("Exclude", "Adaptive data stream classification – general, no payment card focus."),
    1054: ("Exclude", "AIS and ethics – not fraud detection."),
    1056: ("Exclude", "Abnormal behavior detection in virtualized systems – not payment cards."),
    1066: ("Exclude", "Economic anomalies in S&P500 – not fraud detection."),
    1074: ("Exclude", "Marketing in digital business platforms – not fraud detection."),
    1090: ("Exclude", "Multiclass imbalanced learning – general, no payment card focus."),
    1112: ("Exclude", "Deep learning review – mentions fraud detection but no payment card focus."),
    1120: ("Exclude", "Antisocial online behavior detection – not payment card fraud."),
    1159: ("Exclude", "Deep learning for cyber intrusion detection – not financial fraud."),
    1170: ("Exclude", "Crowd anomaly detection – not payment cards."),
    1195: ("Exclude", "Deep learning for risky retail investors – not payment card fraud."),
    1218: ("Exclude", "MTD techniques on cloud – not fraud detection."),
    1264: ("Exclude", "Smart city security review – not payment card fraud."),
    1270: ("Exclude", "Weighted hybrid ensemble for imbalanced data – general, no payment card focus."),
    1276: ("Exclude", "Network intrusion classification – not financial fraud."),
    1283: ("Exclude", "CWGAN-GP for imbalanced data – general, no payment card focus."),
    1319: ("Exclude", "COVID-19 identification in X-ray – not fraud, no payment cards."),
    1321: ("Exclude", "Telemedicine fraud – not payment cards."),
    1351: ("Exclude", "SVM survey – general, no payment card focus."),
    1375: ("Exclude", "Cost-sensitive ensemble for class imbalance – mentions fraud but no payment card focus."),
    1391: ("Exclude", "JNCA bibliometric overview – not fraud detection."),
    1421: ("Exclude", "Feature selection for tax fraud detection – not payment cards."),
}

# ------------------------------------------------------------
# 4. Function to evaluate a row using keyword search
# ------------------------------------------------------------
def evaluate_row(title, abstract):
    text = f"{title} {abstract}".lower()

    has_ai = bool(ai_re.search(text))
    has_fraud = bool(fraud_re.search(text))
    has_card = bool(card_re.search(text))

    if has_ai and has_fraud and has_card:
        return "Include", "Contains AI/ML, fraud/anomaly detection, and payment card terms."
    elif has_ai and has_fraud and not has_card:
        return "Borderline", "Contains AI/ML and fraud terms, but no explicit payment card terminology."
    elif has_ai and has_card and not has_fraud:
        return "Borderline", "Contains AI/ML and payment card terms, but no explicit fraud detection focus."
    elif has_fraud and has_card and not has_ai:
        return "Borderline", "Contains fraud and payment card terms, but no AI/ML methods."
    else:
        return "Exclude", "Does not simultaneously address AI/ML, fraud detection, and payment cards."

# ------------------------------------------------------------
# 5. Apply manual overrides first, then automatic for missing IDs
# ------------------------------------------------------------
def get_decision(row):
    rid = row['id']
    if rid in manual_decisions:
        return manual_decisions[rid]
    else:
        title = str(row.get('A', ''))  # adjust column name if needed
        abstract = str(row.get('abstract', ''))
        return evaluate_row(title, abstract)

# Note: The column names in your Excel might be 'A' for title? Actually from the preview, column A is 'id', B is 'citation(total_reference)', C is 'year', D is 'abstract'. The title is embedded in the citation string. For simplicity, we'll combine the citation and abstract for keyword search.
def get_text(row):
    citation = str(row.get('B', ''))
    abstract = str(row.get('abstract', ''))
    return f"{citation} {abstract}"

# We'll override the evaluate_row to use the combined text
def evaluate_row_combined(text):
    text_lower = text.lower()
    has_ai = bool(ai_re.search(text_lower))
    has_fraud = bool(fraud_re.search(text_lower))
    has_card = bool(card_re.search(text_lower))
    if has_ai and has_fraud and has_card:
        return "Include", "Contains AI/ML, fraud/anomaly detection, and payment card terms."
    elif has_ai and has_fraud and not has_card:
        return "Borderline", "Contains AI/ML and fraud terms, but no explicit payment card terminology."
    elif has_ai and has_card and not has_fraud:
        return "Borderline", "Contains AI/ML and payment card terms, but no explicit fraud detection focus."
    elif has_fraud and has_card and not has_ai:
        return "Borderline", "Contains fraud and payment card terms, but no AI/ML methods."
    else:
        return "Exclude", "Does not simultaneously address AI/ML, fraud detection, and payment cards."

def process_row(row):
    rid = row['id']
    if rid in manual_decisions:
        return manual_decisions[rid]
    else:
        combined = f"{row.get('B', '')} {row.get('abstract', '')}"
        return evaluate_row_combined(combined)

# Apply
df[['Decision', 'Reason']] = df.apply(process_row, axis=1, result_type='expand')

# ------------------------------------------------------------
# 6. Save the annotated file
# ------------------------------------------------------------
output_path = "elsevier_citation_annotated.xlsx"
df.to_excel(output_path, index=False)
print(f"Annotated file saved to {output_path}")
print("\nSummary of decisions:")
print(df['Decision'].value_counts())
